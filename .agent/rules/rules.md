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
- **README.md Lifecycle**: Update and maintain `@README.md` immediately following successful feature additions. 
  - **Development Versions**: Dev version notes may be actively maintained during iterative work cycles.
  - **Stable Releases**: If a final stable release is requested, strip all dev-version documentation, structure the document cleanly, and explicitly compare updated features against the previous stable baseline.
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
- **README.md Lifecycle**: Update and maintain `@README.md` immediately following successful feature additions. 
  - **Development Versions**: Dev version notes may be actively maintained during iterative work cycles.
  - **Stable Releases**: If a final stable release is requested, strip all dev-version documentation, structure the document cleanly, and explicitly compare updated features against the previous stable baseline.

## 7. Strict Deployment & Terminal Command Controls
- **No Automatic Git Pushes**: You are strictly forbidden from running `git push` commands to GitHub or any remote repository. 
- **Local Commits Only**: You may only stage or commit files locally (`git add`, `git commit`).
- **Dev Branch Archiving**: All development version commits must be isolated and committed directly to the `archive-dev-BOA` local branch. The `main` branch must only receive stable commits.
- **Explicit Permission Gate**: A feature must be explicitly greenlit by the user before it is marked as finalized, merged into the local `archive-dev-BOA` branch, or pushed to `main`.

## 8. GUI Architecture & Navigation Rules
- **Explicit Accelerator Keys**: Every control inside a WxPython or NVDA Settings panel must have a unique, explicit accelerator key assigned via the ampersand symbol (`&`). Never rely on Windows' default first-letter navigation. When attempting to assign shortcuts to Group Boxes (`wx.StaticBox`), always apply the `&` to the **first child control** within the group, as Windows natively ignores accelerators on StaticBox titles inside complex embedded dialogs.