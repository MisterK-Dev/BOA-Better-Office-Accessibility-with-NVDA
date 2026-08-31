# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

"""
NVDA API Compatibility Database Updater.
This script parses NVDA Developer Guides and Changelogs (both local files and official web links)
to identify deprecated or removed APIs, and interactively updates the compatibility matrix JSON file.
"""

import os
import sys
import ast
import json
import re
import shutil
import urllib.request
import urllib.parse
from html.parser import HTMLParser

# --- HTML Parser for NVDA Changes Document ---

class NVDAChangesParser(HTMLParser):
    """
    Parses NVDA changes.html (or developerGuide.html) to extract developer changes
    and API breaking changes for a target version.
    """
    def __init__(self, target_version):
        super().__init__()
        # Target version is normalized to major.minor (e.g. '2026.1')
        self.target_version = ".".join(target_version.split(".")[:2])
        self.in_target_version = False
        self.in_dev_section = False
        self.in_li = False
        self.current_li_text = []
        self.current_li_codes = []
        self.current_tag = None
        self.version_changes = []

    def handle_starttag(self, tag, attrs):
        self.current_tag = tag
        if tag == "li" and (self.in_target_version or self.target_version == "force"):
            if self.in_dev_section or self.target_version == "force":
                self.in_li = True
                self.current_li_text = []
                self.current_li_codes = []

    def handle_endtag(self, tag):
        if tag == "li" and self.in_li:
            self.in_li = False
            full_text = "".join(self.current_li_text).strip()
            full_text = " ".join(full_text.split())
            if full_text:
                self.version_changes.append({
                    "text": full_text,
                    "codes": list(self.current_li_codes)
                })
        self.current_tag = None

    def handle_data(self, data):
        if self.current_tag == "h2":
            val = data.strip()
            # Match version numbers like '2026.1'
            m = re.search(r"\d+\.\d+", val)
            if m:
                ver = m.group(0)
                if ver == self.target_version:
                    self.in_target_version = True
                else:
                    self.in_target_version = False
                    self.in_dev_section = False
            else:
                self.in_target_version = False
                self.in_dev_section = False
                
        elif self.current_tag in ("h3", "h4") and self.in_target_version:
            val = data.strip().lower()
            if "changes for developers" in val or "api breaking changes" in val or "deprecations" in val:
                self.in_dev_section = True
            else:
                # Keep dev section active unless another major heading is encountered
                pass
                
        elif self.in_li:
            self.current_li_text.append(data)
            if self.current_tag == "code":
                cleaned_code = data.strip()
                # Exclude basic python types/words to prevent pollution
                if cleaned_code and cleaned_code not in self.current_li_codes:
                    self.current_li_codes.append(cleaned_code)


# --- Helper Functions ---

def get_target_versions(buildvars_path):
    """Statically parses buildVars.py to extract target NVDA versions."""
    with open(buildvars_path, "r", encoding="utf-8") as f:
        tree = ast.parse(f.read())
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name) and target.id == "addon_info":
                        if isinstance(node.value, ast.Dict):
                            min_ver = None
                            last_tested = None
                            for k, v in zip(node.value.keys, node.value.values):
                                k_str = k.value if isinstance(k, ast.Constant) else None
                                v_str = v.value if isinstance(v, ast.Constant) else None
                                if k_str == "addon_minimumNVDAVersion":
                                    min_ver = v_str
                                elif k_str == "addon_lastTestedNVDAVersion":
                                    last_tested = v_str
                            return min_ver, last_tested
    return None, None


def verify_reference_documentation(refs_dir, last_tested_version):
    """Checks if there is an official reference document on disk in refs_dir containing version string or its major.minor prefix."""
    if not os.path.exists(refs_dir):
        return False
    candidates = [last_tested_version]
    two_part_ver = ".".join(last_tested_version.split(".")[:2])
    if two_part_ver not in candidates:
        candidates.append(two_part_ver)
    for file in os.listdir(refs_dir):
        for cand in candidates:
            if cand in file:
                return True
    return False


def decode_mhtml_to_html(filepath):
    """Decodes MHTML (multipart/related, quoted-printable) to raw HTML if needed."""
    if filepath.endswith(".mhtml"):
        import email
        from email import policy
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            msg = email.message_from_file(f, policy=policy.default)
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                return part.get_content()
    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()


def analyze_change_text(item, target_version, citation):
    """
    Analyzes raw item text using heuristics to extract the deprecated/removed symbol,
    its status, and recommended alternative.
    """
    text = item["text"]
    codes = item["codes"]
    if not codes:
        return None

    text_lower = text.lower()
    is_removed = any(x in text_lower for x in ("removed", "no longer supported", "has been removed", "have been removed"))
    is_deprecated = any(x in text_lower for x in ("deprecated", "deprecation"))
    is_moved_or_renamed = any(x in text_lower for x in ("moved", "renamed", "replaced"))

    if not (is_removed or is_deprecated or is_moved_or_renamed):
        return None

    status = "removed" if is_removed else "deprecated"
    symbol = codes[0]

    # Filter out generic python keywords/types
    if symbol in ("bool", "int", "None", "True", "False", "dict", "list", "set", "str", "tuple"):
        if len(codes) > 1:
            symbol = codes[1]
            remaining_codes = codes[2:]
        else:
            return None
    else:
        remaining_codes = codes[1:]

    alternative = "no replacement"
    if "no replacement" in text_lower or "no direct replacement" in text_lower:
        alternative = "no replacement"
    elif remaining_codes:
        alternative = remaining_codes[0]

    return {
        "symbol": symbol,
        "status": status,
        "version": target_version,
        "alternative": alternative,
        "citation": citation
    }


def download_url(url, target_path):
    """Downloads a file from a URL to target_path using urllib."""
    req = urllib.request.Request(
        url,
        headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    )
    with urllib.request.urlopen(req) as response_stream:
        with open(target_path, "wb") as out_file:
            out_file.write(response_stream.read())


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "..", ".."))
    
    buildvars_path = os.path.join(root_dir, "buildVars.py")
    matrix_path = os.path.join(script_dir, "..", "api_matrix.json")
    refs_dir = os.path.join(root_dir, ".agent", "REFS_Sources")

    print("="*60)
    print("NVDA API COMPATIBILITY DATABASE UPDATER (MANUAL UTILITY)")
    print("="*60)

    # 1. Read buildVars.py versions
    min_ver, last_tested = get_target_versions(buildvars_path)
    if not min_ver or not last_tested:
        print("[ERROR] Could not extract target NVDA versions from buildVars.py")
        sys.exit(1)

    print(f"[*] Read target NVDA version from buildVars.py: {last_tested}")

    # 2. Check if reference document exists. If missing, ask the user!
    has_ref = verify_reference_documentation(refs_dir, last_tested)
    if not has_ref:
        print(f"\n[!] Missing latest reference document for '{last_tested}' in .agent/REFS_Sources/")
        while True:
            response = input("Please specify a local file path, a download URL, or 'skip': ").strip()
            if not response:
                continue
            if response.lower() == 'skip':
                print("[*] Skipping local reference checks. Web-only links will be utilized.")
                break
            elif response.startswith("http://") or response.startswith("https://"):
                print(f"[*] Downloading reference from URL: {response}")
                try:
                    os.makedirs(refs_dir, exist_ok=True)
                    url_path = urllib.parse.urlparse(response).path
                    filename = os.path.basename(url_path) or f"nvda_{last_tested}_ref.html"
                    target_path = os.path.join(refs_dir, filename)
                    download_url(response, target_path)
                    print(f"[SUCCESS] Saved to {target_path}")
                    break
                except Exception as e:
                    print(f"[ERROR] Failed to download URL: {e}")
            else:
                if os.path.exists(response):
                    os.makedirs(refs_dir, exist_ok=True)
                    filename = os.path.basename(response)
                    target_path = os.path.join(refs_dir, filename)
                    shutil.copy2(response, target_path)
                    print(f"[SUCCESS] Copied file to {target_path}")
                    break
                else:
                    print(f"[ERROR] Path does not exist: {response}. Please retry.")

    # 3. Compile sources to parse (Local and Web Links)
    sources = []
    
    # Add local files from .agent/REFS_Sources
    if os.path.exists(refs_dir):
        for file in os.listdir(refs_dir):
            if last_tested in file or "Developer Guide" in file:
                sources.append({
                    "type": "local",
                    "path": os.path.join(refs_dir, file),
                    "citation": f"Local file: {file}"
                })

    # Add official web links
    web_links = [
        "https://www.nvaccess.org/files/nvda/documentation/changes.html",
        "https://download.nvaccess.org/documentation/developerGuide.html"
    ]
    for link in web_links:
        sources.append({
            "type": "web",
            "url": link,
            "citation": link
        })

    # 4. Parse all sources
    proposed_updates = []
    
    for src in sources:
        try:
            if src["type"] == "local":
                print(f"[*] Reading and parsing local file: {src['path']}")
                html_content = decode_mhtml_to_html(src["path"])
                parser = NVDAChangesParser(last_tested)
                parser.feed(html_content)
                
                # Fallback if no version header was matched (treat entire file as target version)
                if not parser.version_changes:
                    parser = NVDAChangesParser("force")
                    parser.feed(html_content)
            else:
                print(f"[*] Fetching and parsing web link: {src['url']}")
                req = urllib.request.Request(
                    src["url"],
                    headers={'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
                )
                with urllib.request.urlopen(req, timeout=10) as resp:
                    html_content = resp.read().decode('utf-8', errors='ignore')
                parser = NVDAChangesParser(last_tested)
                parser.feed(html_content)

            # Analyze changes and build proposed list
            for item in parser.version_changes:
                analysis = analyze_change_text(item, last_tested, src["citation"])
                if analysis:
                    # Avoid adding exact duplicates in the proposed list
                    if not any(x["symbol"] == analysis["symbol"] for x in proposed_updates):
                        proposed_updates.append(analysis)
        except Exception as e:
            print(f"[WARNING] Could not parse source {src.get('path') or src.get('url')}: {e}")

    # 5. Load existing database
    existing_matrix = {}
    if os.path.exists(matrix_path):
        with open(matrix_path, "r", encoding="utf-8") as f:
            try:
                existing_matrix = json.load(f)
            except Exception as e:
                print(f"[WARNING] Failed to parse existing {matrix_path}: {e}")

    # 6. Interactive Approval Loop
    if not proposed_updates:
        print("\n[*] No API deprecations or removals discovered from the scanned documentation.")
        sys.exit(0)

    print(f"\n[*] Discovered {len(proposed_updates)} potential API compatibility constraints.")
    print("Please review each candidate carefully. You can accept, reject, or edit them.")
    print("-" * 60)

    updates_made = False

    for update in proposed_updates:
        symbol = update["symbol"]
        status = update["status"]
        alt = update["alternative"]
        citation = update["citation"]
        
        # Check if already present with same values
        if symbol in existing_matrix:
            curr = existing_matrix[symbol]
            if curr.get("status") == status and curr.get("alternative") == alt:
                # Already up to date, skip asking
                continue

        print(f"\nProposed API Entry:")
        print(f"  -> Symbol:      {symbol}")
        print(f"  -> Status:      {status}")
        print(f"  -> Alternative: {alt}")
        print(f"  -> Citation:    {citation}")
        
        while True:
            choice = input("Add/Update this entry? [y]es / [n]o / [e]dit / [q]uit: ").strip().lower()
            if choice == 'y':
                existing_matrix[symbol] = {
                    "status": status,
                    "version": last_tested,
                    "alternative": alt,
                    "citation": citation
                }
                updates_made = True
                print("[*] Accepted entry.")
                break
            elif choice == 'n':
                print("[*] Skipped entry.")
                break
            elif choice == 'e':
                # Allow manual overrides to prevent any AI parsing errors
                new_symbol = input(f"Enter symbol [{symbol}]: ").strip() or symbol
                new_status = input(f"Enter status [{status}]: ").strip() or status
                new_alt = input(f"Enter alternative [{alt}]: ").strip() or alt
                existing_matrix[new_symbol] = {
                    "status": new_status,
                    "version": last_tested,
                    "alternative": new_alt,
                    "citation": citation
                }
                updates_made = True
                print("[*] Saved custom entry.")
                break
            elif choice == 'q':
                print("[*] Exiting updater process.")
                break
        if choice == 'q':
            break

    # 7. Write back to database if changes approved
    if updates_made:
        with open(matrix_path, "w", encoding="utf-8") as f:
            json.dump(existing_matrix, f, indent=2, ensure_ascii=False)
        print(f"\n[SUCCESS] Compatibility matrix database successfully saved: {matrix_path}")
    else:
        print("\n[*] No modifications were written to the database.")

if __name__ == "__main__":
    main()
