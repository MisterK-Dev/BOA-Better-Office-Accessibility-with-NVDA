---
description: "Global quality, safety, and Git rules for NVDA Addon Core development"
alwaysApply: true
---

# Antigravity Global Project Rules: NVDA Addon Core

## 1. Context Retention & Workspace Orientation
- **Pre-Execution Scanning**: Before writing code, you must thoroughly scan the active workspace, folder structure, core file imports, and dependencies.
- **Never Forget Past Context**: Review previous feature discussions, structural choices, and historical design patterns before implementing new blocks. you can take previous conversations, and related details when ever you requires.
- **Learn From Mistakes**: Maintain an internal trace of past architectural bugs, clashing keystrokes, and type errors encountered in this repository. Cross-reference them to avoid repetition.
- **Continuous Evolution**: Treat new feature requests as increments of the existing architecture. Do not suggest rewrites or radical refactors unless explicitly asked.

## 2. NVDA Safety & Performance Constraints
- **Zero Harm to Core Behavior**: New features must never degrade standard NVDA operations, native speech queues, existing features, or global focus tracking.
- **Single-Threaded Safety**: NVDA is single-threaded. Never execute standard Python blocking loops (`time.sleep`). All delays or async checks must use `tonictimer.wxTimer` or `queueHandler.queueFunction`.
- **Output Routing**: Route all user-facing speech and braille notifications cleanly through `ui.message`. Never use raw `print()` statements for core announcements.
- **Modular Isolation**: Keep feature code out of `globalPlugin.py`. Isolate new logic inside independent local modules and use clean wrapper imports.

## 3. Mandatory Pre-Finalization Verification
Before outputting any finalized code blocks, scripts, or diff artifacts in the workspace view, execute a strict triple-pass validation:
1. **Syntax & Typo Check**: Verify full Python 3 compliance, check indentation, find missing colons, and audit variable spelling.
2. **Logical Path Verification**: Trace the execution flow to confirm variable scopes do not break across NVDA's dynamic event loops.
3. **API Integrity**: Double-check that all referenced NVDA internal module methods (like `speech`, `braille`, `controlTypes`) match actual NVDA API specifications.

## 4. Keystroke Conflict Check
- **Gesture Audit**: Before generating or suggesting any new shortcut gesture (e.g., `kb:nvda+control+f`), perform a comprehensive scan across all existing script files and configuration blocks.
- **Conflict Prevention**: Verify that the proposed shortcut is completely unique and is not already mapped to an active or deprecated feature within this addon.

## 5. Security & Vulnerability Controls
- **Secure Code Quality**: All generated code must be securely structured, completely free from common vulnerabilities, and follow safe input/output handling.

## 6. Automated Docstring, Commenting & README Maintenance
- **Intent-Driven Docstrings**: Every single function, class, or method written must include an explicit, detailed docstring explaining the architectural *why* behind the logic.
- **Granular Code Comments**: Every logical step, complex condition, or NVDA API interaction must be accompanied by detailed line comments.
- **README.md will be mentained by the developer itself from now onwards no need to touch it.**: [no need to update and mentain`@README.md`] immediately following successful feature additions. 

## 7. Strict Deployment & Terminal Command Controls
- **No Automatic Git Pushes**: You are strictly forbidden from running `git push` commands to GitHub or any remote repository automatically. You may only push to remote branches if explicitly directed, reviewed, and approved by the user.
- **Local Commits Only**: You may only stage or commit files locally (`git add`, `git commit`) unless granted the final greenlight to push.
- **Dev Branch Archiving**: All development version commits must be isolated and committed directly to the `archive-dev-BOA` local branch. The `main` branch must only receive stable commits.
- **Explicit Permission Gate**: A feature must be explicitly greenlit by the user before it is marked as finalized, merged into the local `archive-dev-BOA` branch, or pushed to `main`. Before any push, you MUST execute the Release Validation Pipeline.

## 8. GUI Architecture & Navigation Rules
- **Explicit Accelerator Keys**: Every control inside a WxPython or NVDA Settings panel must have a unique, explicit accelerator key assigned via the ampersand symbol (`&`). Never rely on Windows' default first-letter navigation. When attempting to assign shortcuts to Group Boxes (`wx.StaticBox`), always apply the `&` to the **first child control** within the group, as Windows natively ignores accelerators on StaticBox titles inside complex embedded dialogs.

## 9. Excel COM Object Safety
- **Anti-Freeze Bailouts**: Never evaluate, loop over, or request properties from millions of Excel cells simultaneously across the COM bridge. COM marshaling between Excel and NVDA is extremely slow. Always implement a strict safety bailout threshold (e.g., `if count > 2000: break`) when linearly scanning Excel structures to prevent catastrophic NVDA freezes.

## 10. Absolute Layout Boundaries
- **Edge Detection**: When performing layout analysis or structural bounds checking (e.g., "Are the edge columns hidden?"), **never rely on `UsedRange`**. `UsedRange` ignores manually hidden rows/columns that lie far outside the active data. Always check the absolute mathematical boundaries of the grid (Row `1048576`, Column `16384`).

## 11. Scripting & File Encoding Traps
- **No PowerShell Appending**: Never use native PowerShell output redirection (`>>` or `>`) to append text to codebase files. PowerShell defaults to UTF-16 LE encoding, which corrupts standard UTF-8 files by injecting invisible null-bytes. Always use native Python or strict UTF-8 file editing tools to manipulate repository documentation.

## 12. Office COM Exception Handling
- **Graceful Failure**: The Microsoft Office COM bridge is inherently unstable. If Excel is in "Cell Editing Mode", a dialog box is open, or the application is busy, simple COM calls (like fetching `ActiveSheet`) will violently throw `com_error` or `HRESULT` exceptions. *Always* wrap direct Office COM queries in a `try...except Exception:` block and fail gracefully. Never allow a raw COM exception to bubble up and crash NVDA.

## 13. Safe Event Hooking (`nextHandler`)
- **Event Forwarding**: When hooking into global NVDA events (like `event_NVDAObject_init` or `event_gainFocus`), you must *always* execute the fallback `nextHandler()` or `super()` method at the end of your logic. Failing to do so will intercept the event completely, effectively blinding NVDA and permanently breaking its ability to read standard Windows objects.

## 14 Internationalization (i18n) & Translation Governance
1. **Strict Wrapping:** Any new user-facing English strings (UI labels, dialogs, speech announcements) MUST be wrapped in the `_()` translation function.
2. **No F-Strings:** You MUST NEVER use Python f-strings for user-facing text (e.g., `_("Row {r}")`). You must use `.format()` structures (e.g., `_("Row {r}").format(r=row_num)`) to ensure translation tools can extract the template cleanly.
3. **Auto-POT Generation:** Upon completing any new feature or modifying existing user-facing strings, you MUST automatically run the `scons pot` command to regenerate the master `BOA.pot` template.
4. **Agent Auto-Translation Skill:** After generating the `.pot` file, if the user requests translation, you must execute the multi-lingual translation pipeline. Refer to `.agent/skills/translation_pipeline/translation_pipeline.md` for the exact step-by-step instructions.

## 15. Copyright & Licensing
- **Mandatory Headers**: Every new Python (`.py`) file created or modified for this project MUST include the standard copyright header at the very top of the file:
```python
# -*- coding: UTF-8 -*-
# Copyright (C) 2026 KIRAN G T {MisterK} and Antigravity 2
# This file is covered by the GNU General Public License (GPL), version 2.
# See the file COPYING.txt for more details.
```

## 16. Branch Syncing Architecture
- **Syncing Stable to Dev**: When syncing release updates, large translations, or finalized tags from `main` back into the `archive-dev-BOA` branch, you must **always use a Squash Merge** (`git merge --squash main`). This keeps the archive history perfectly clean as a standalone log without inheriting `main`'s granular commit history.
