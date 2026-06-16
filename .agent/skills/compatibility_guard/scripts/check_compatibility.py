# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.

import os
import sys
import ast
import json

def resolve_attribute_path(node, imported_modules):
    """Recursively resolves attribute access chains (e.g. speech.speech.IDT_TONE_DURATION) using imports."""
    if isinstance(node, ast.Name):
        return imported_modules.get(node.id, node.id)
    elif isinstance(node, ast.Attribute):
        prefix = resolve_attribute_path(node.value, imported_modules)
        if prefix:
            return f"{prefix}.{node.attr}"
    return None

class CompatibilityVisitor(ast.NodeVisitor):
    def __init__(self, filepath, blacklist):
        self.filepath = filepath
        self.blacklist = blacklist
        self.imported_modules = {}  # alias -> full module name
        self.imported_names = {}    # name -> full symbol path
        self.errors = []

    def visit_Import(self, node):
        for alias in node.names:
            self.imported_modules[alias.asname or alias.name] = alias.name
        self.generic_visit(node)

    def visit_ImportFrom(self, node):
        module_name = node.module
        if module_name:
            for alias in node.names:
                full_name = f"{module_name}.{alias.name}"
                self.imported_names[alias.asname or alias.name] = full_name
                # Check direct imports
                if full_name in self.blacklist:
                    self.errors.append((node.lineno, f"Direct import of blacklisted/deprecated symbol '{full_name}'"))
        self.generic_visit(node)

    def visit_Attribute(self, node):
        # Resolve deep attribute paths, e.g. speech.speech.IDT_TONE_DURATION
        full_path = resolve_attribute_path(node, self.imported_modules)
        if full_path and full_path in self.blacklist:
            self.errors.append((node.lineno, f"Usage of blacklisted/deprecated symbol '{full_path}'"))
        self.generic_visit(node)

    def visit_Name(self, node):
        # E.g. from speech import speakText; speakText("foo")
        if node.id in self.imported_names:
            full_path = self.imported_names[node.id]
            if full_path in self.blacklist:
                self.errors.append((node.lineno, f"Usage of blacklisted/deprecated symbol '{full_path}'"))
        self.generic_visit(node)

def get_target_versions(buildvars_path):
    """Statically parses buildVars.py to extract NVDA version constraints."""
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
                                k_str = None
                                v_str = None
                                if isinstance(k, ast.Constant):
                                    k_str = k.value
                                if isinstance(v, ast.Constant):
                                    v_str = v.value
                                
                                if k_str == "addon_minimumNVDAVersion":
                                    min_ver = v_str
                                elif k_str == "addon_lastTestedNVDAVersion":
                                    last_tested = v_str
                            return min_ver, last_tested
    return None, None

def verify_reference_documentation(refs_dir, last_tested_version):
    """
    Checks if there is an official reference document on disk in .agent/REFS_Sources/
    containing the last tested version string to prevent AI hallucinations.
    """
    if not os.path.exists(refs_dir):
        return False
    
    # Check if any file in the references directory contains the version string
    for file in os.listdir(refs_dir):
        if last_tested_version in file:
            return True
    return False

def scan_codebase(addon_dir, blacklist):
    all_errors = {}
    for root, dirs, files in os.walk(addon_dir):
        if "__pycache__" in root or ".sconsign" in root:
            continue
        for file in files:
            if file.endswith(".py"):
                filepath = os.path.join(root, file)
                with open(filepath, "r", encoding="utf-8") as f:
                    try:
                        tree = ast.parse(f.read())
                        visitor = CompatibilityVisitor(filepath, blacklist)
                        visitor.visit(tree)
                        if visitor.errors:
                            all_errors[filepath] = visitor.errors
                    except Exception as e:
                        print(f"Error parsing {filepath}: {e}")
    return all_errors

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    root_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "..", ".."))
    
    buildvars_path = os.path.join(root_dir, "buildVars.py")
    matrix_path = os.path.join(script_dir, "..", "api_matrix.json")
    refs_dir = os.path.join(root_dir, ".agent", "REFS_Sources")
    addon_dir = os.path.join(root_dir, "addon")

    print(f"Loading version constraints from: {buildvars_path}")
    min_ver, last_tested = get_target_versions(buildvars_path)
    
    if not min_ver or not last_tested:
        print("Error: Could not extract target NVDA versions from buildVars.py")
        sys.exit(1)

    print(f" -> Minimum Supported: {min_ver}")
    print(f" -> Last Tested/Target: {last_tested}")

    # 1. Gating check: Ensure authentic reference documentation exists for the target version
    print(f"Verifying reference documentation in: {refs_dir}")
    if not verify_reference_documentation(refs_dir, last_tested):
        print(f"\n[WARNING] Version Gating Triggered!")
        print(f"No official reference files containing '{last_tested}' found in .agent/REFS_Sources/.")
        print(f"To ensure compatibility verification is grounded and to prevent AI hallucinations,")
        print(f"please download/save the official NVDA Developer Guide or Changelog for {last_tested}")
        print(f"into .agent/REFS_Sources/ before finalizing.")
        print("-" * 50)

    # 2. Load the blacklisted symbols
    if not os.path.exists(matrix_path):
        print(f"Error: Compatibility matrix file not found at {matrix_path}")
        sys.exit(1)
        
    with open(matrix_path, "r", encoding="utf-8") as f:
        blacklist = json.load(f)

    # 3. Scan the codebase
    print(f"Scanning BOA codebase for API compatibility violations in: {addon_dir}")
    all_errors = scan_codebase(addon_dir, blacklist)

    print("\n" + "="*50)
    print("BOA NVDA API COMPATIBILITY AUDIT REPORT")
    print("="*50)

    if all_errors:
        print(f"\nDiscovered {sum(len(errs) for errs in all_errors.values())} API violations:")
        for filepath, errors in sorted(all_errors.items()):
            rel_path = os.path.relpath(filepath, root_dir)
            print(f"\nFile: {rel_path}")
            for line, msg in errors:
                print(f"  Line {line}: {msg}")
                # Print citation and alternative recommendation
                symbol = msg.split("'")[1]
                if symbol in blacklist:
                    info = blacklist[symbol]
                    print(f"    -> Status: {info['status']} in version {info['version']}")
                    print(f"    -> Alternative: Use '{info['alternative']}' instead")
                    print(f"    -> Citation source: {info['citation']}")
        
        print("\nAudit Failed! Please resolve the compatibility violations above.")
        sys.exit(1)
    else:
        print("\nAudit Passed! No deprecated or removed NVDA API usages discovered.")
        sys.exit(0)

if __name__ == "__main__":
    main()
