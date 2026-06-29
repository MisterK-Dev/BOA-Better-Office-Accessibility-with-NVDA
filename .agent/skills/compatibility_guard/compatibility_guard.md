# Agent Skill: NVDA Version Compatibility Guard

**Purpose:**
This skill acts as a static compatibility checker for the BOA codebase. It ensures we do not accidentally import deprecated or removed NVDA APIs, preventing runtime crashes across our target NVDA versions.

---

## Target Version Scope

* **Minimum Supported Version**: `2026.1.0`
* **Last Tested Version**: `2026.1.1`

Since we target the modern `2026.1.x` release range, we must focus strictly on preventing the usage of legacy/deprecated code snippets that were active in older NVDA versions (pre-2026.1) but are now removed.

---

## Unskippable Sources of Truth

Every compatibility rule and blacklisted symbol in `api_matrix.json` must be verified and cited directly from the following official documentation pages:
1. **NV Access Developer Guide**: `https://download.nvaccess.org/documentation/developerGuide.html`
2. **NVDA Add-on Development Guide Wiki**: `https://github.com/nvdaaddons/DevGuide/wiki/NVDA-Add-on-Development-Guide`

No API rules may be added to `api_matrix.json` without an active `"citation"` linking to these official pages.

---

## How to Execute the Compatibility Guard

To check the codebase for compatibility violations, run the following command from the root of the repository:

```powershell
python .agent/skills/compatibility_guard/scripts/check_compatibility.py
```

### Future Updates & Database Maintenance

Whenever you bump the target version in `buildVars.py` (e.g. to a future `2026.2` release), you can automatically update the compatibility matrix database.

To run the interactive database updater, execute the following command:

```powershell
python .agent/skills/compatibility_guard/scripts/update_matrix.py
```

This script will:
1. Parse `buildVars.py` for target versions.
2. Verify that the corresponding developer guide or changes reference document is present in `.agent/REFS_Sources/`. If the document is missing, the script will halt and prompt you to specify a local path or a download URL.
3. Automatically fetch and parse deprecations/removals from both the local reference document and the online NVDA changes website.
4. Present each proposed update interactively, allowing you to approve, skip, or manually edit the entry before writing it to `api_matrix.json` (ensuring 100% accuracy and zero hallucination).
