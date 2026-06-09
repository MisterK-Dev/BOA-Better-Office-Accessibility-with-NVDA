# Antigravity Workspace Rules: NVDA Addon Core (BOA)

> [Scope: Active Workspace Root | Priority: High | Activation: Always On]

## 1. Context Retention & Comprehensive Workspace Orientation
- **Deep Codebase Mapping**: Before proposing edits, writing features, or modifying code, you must execute a comprehensive analysis of the entire repository structure. You must actively trace the directory tree, inspect the core entry points (e.g., `globalPlugin.py` or local entry wrappers), and map internal module dependencies to achieve an absolute understanding of the codebase architecture.
- **Historical Alignment**: Review the last 5 local git commit messages (`git log -n 5`) and cross-reference the project layout to align new logic with historical patterns.
- **No Speculative Refactoring**: Treat new feature requests as increments of the existing architecture. Do not suggest or execute structural rewrites or radical refactors unless explicitly ordered.

## 2. Self-Documenting Memory Subsystem (`post-mortem.md`)
- **Mandatory Memory Dump**: You must immediately update or create [create if it not exists] a `post-mortem.md` file at the project root whenever you:
  1. Encounter and resolve a critical Python, WxPython, or NVDA API bug.
  2. Complete a major feature milestone or alter a global state mechanism.
  3. Change project configurations or user-interaction preferences.
- **File Structure Constraint**: Maintain `post-mortem.md` using the following exact structure:
  - `# Architectural Choices & Preferences`: Core configurations or user design choices.
  - `# Active Mistakes & Bug Mitigations`: A running list of discovered traps, clashing gestures, or API workarounds. (Remove or move to `# Resolved Traps` once a permanent fix is verified).
- **Size Optimization**: Keep the file concise. Group similar errors together and limit total file length to under 300 lines to prevent context bloat.

## 3. Strict Temporary File Lifecycle & Workspace Cleanup
- **Ephemeral Testing Isolation**: Any script, log dump, raw text snippet, or mock data file created solely for intermediate testing or verification must be tracked continuously during your session.
- **Automated Mandatory Cleanup**: Prior to finalization, marking a task complete, or prompting the user for verification, you must delete every single temporary file you generated that serves no functional runtime purpose.
- **Safety Boundaries for Deletion**: 
  1. You are strictly forbidden from deleting any pre-existing file in the workspace or any file tracked by Git unless explicitly requested.
  2. Run `git status --ignored` or check tracking status before executing file deletions (`os.remove` or terminal removals) to guarantee no source assets or structural dependencies are targeted.

## 4. NVDA Safety & Performance Constraints
- **Zero Harm to Core Behavior**: New features must never degrade standard NVDA operations, native speech queues, existing features, or global focus tracking.
- **Single-Threaded Safety**: NVDA is single-threaded. Never execute standard Python blocking loops (`time.sleep`). All delays or async non-blocking loops must use `core.callLater(ms, func)`. Use `queueHandler.queueFunction(queueHandler.eventQueue, func)` strictly for when a background thread needs to push data back to NVDA's main thread.
- **Output Routing**: Route all user-facing speech and braille notifications cleanly through `ui.message`. Never use raw `print()` statements for core announcements.
- **Modular Isolation**: Keep feature code out of `globalPlugin.py`. Isolate new logic inside independent local modules and use clean wrapper imports.

## 5. Mandatory Pre-Finalization Verification
Before outputting any finalized code blocks, scripts, or diff artifacts, execute a strict triple-pass validation:
1. **Syntax & Typo Check**: Verify full Python 3 compliance, check indentation, find missing colons, and audit variable spelling.
2. **Logical Path Verification**: Trace the execution flow to confirm variable scopes do not break across NVDA's dynamic event loops.
3. **API Integrity**: Double-check that all referenced NVDA internal module methods (like `speech`, `braille`, `controlTypes`) match actual NVDA API specifications.

## 6. Keystroke Conflict Check
- **Gesture Audit**: Before generating or suggesting any new shortcut gesture (e.g., `kb:nvda+control+f`), perform a comprehensive scan across all existing script files and configuration blocks.
- **Conflict Prevention**: Verify that the proposed shortcut is completely unique and is not already mapped to an active or deprecated feature within this addon.

## 7. Security & Vulnerability Controls
- **Secure Code Quality**: All generated code must be securely structured, completely free from common vulnerabilities, and follow safe input/output handling.

## 8. Automated Docstring, Commenting & README Maintenance
- **Intent-Driven Docstrings**: Every single function, class, or method written must include an explicit, detailed docstring explaining the architectural *why* behind the logic.
- **Granular Code Comments**: Every logical step, complex condition, or NVDA API interaction must be accompanied by detailed line comments.
- **README.md Lifecycle**: Update and maintain `README.md` immediately following successful feature additions. 
  - **Development Versions**: Dev version notes may be actively maintained during iterative work cycles.
  - **Stable Releases**: If a final stable release is requested, strip all dev-version documentation, structure the document cleanly, and explicitly compare updated features against the previous stable baseline.

## 9. Strict Deployment & Terminal Command Controls
- **No Automatic Git Pushes**: You are strictly forbidden from running `git push` commands to GitHub or any remote repository. 
- **Local Commits Only**: You may only stage or commit files locally (`git add`, `git commit`).
- **Dev Branch Archiving**: All development version commits must be isolated and committed directly to the `archive-dev-BOA` local branch. The `main` branch must only receive stable commits.
- **Explicit Permission Gate**: A feature must be explicitly greenlit by the user before it is marked as finalized, merged into the local `archive-dev-BOA` branch, or pushed to `main`.

## 10. GUI Architecture & Navigation Rules
- **Explicit Accelerator Keys**: Strictly require `&` accelerators for controls inside NVDA Settings Panels (`NVDASettingsDialog`). Custom, ephemeral pop-up dialogs (like `wx.TextEntryDialog` or Layout reports) are exempt to prevent UI clutter. Never rely on Windows' default first-letter navigation. When attempting to assign shortcuts to Group Boxes (`wx.StaticBox`) inside settings, always apply the `&` to the **first child control** within the group, as Windows natively ignores accelerators on StaticBox titles inside complex embedded dialogs.

## 11. Excel COM Object Safety
- **Anti-Freeze Bailouts & SpecialCells**: Never evaluate or loop over millions of Excel cells simultaneously across the COM bridge. COM marshaling is extremely slow. Whenever possible, avoid Python COM loops and use Excel's native `Range.SpecialCells()` (e.g., `xlCellTypeVisible`) to offload logic to the internal C++ engine. If you must loop linearly, always implement a strict safety bailout threshold (e.g., `if count > 2000: break`) to prevent catastrophic freezes.

## 12. Absolute Layout Boundaries
- **Edge Detection**: When performing layout analysis or structural bounds checking (e.g., "Are the edge columns hidden?"), **never rely on `UsedRange`**. `UsedRange` ignores manually hidden rows/columns that lie far outside the active data. Always check the absolute mathematical boundaries of the grid (Row `1048576`, Column `16384`).

## 13. Scripting & File Encoding Traps
- **No PowerShell Appending**: Never use native PowerShell output redirection (`>>` or `>`) to append text to codebase files. PowerShell defaults to UTF-16 LE encoding, which corrupts standard UTF-8 files by injecting invisible null-bytes. Always use native Python or strict UTF-8 file editing tools to manipulate repository documentation.

## 14. Office COM Exception Handling
- **Graceful Failure**: The Microsoft Office COM bridge is inherently unstable. If Excel is in "Cell Editing Mode", a dialog box is open, or the application is busy, simple COM calls (like fetching `ActiveSheet`) will violently throw `com_error` or `HRESULT` exceptions. *Always* wrap direct Office COM queries in a `try...except Exception:` block and fail gracefully. Never allow a raw COM exception to bubble up and crash NVDA.

## 15. Safe Event Hooking (`nextHandler`)
- **Event Forwarding**: When hooking into global NVDA events (like `event_NVDAObject_init` or `event_gainFocus`), you must *always* execute the fallback `nextHandler()` or `super()` method at the end of your logic. Failing to do so will intercept the event completely, effectively blinding NVDA and permanently breaking its ability to read standard Windows objects.

## 16. Dynamic MRO (Method Resolution Order) Law
- **Mixins & Super()**: When writing NVDA object overlays, remember that classes are dynamically injected into `clsList` at runtime. Calling `super()` does NOT call a static lexical parent class; it calls the *next* class in the dynamic inheritance chain. Never assume `super()` maps to your class's lexical parent.

## 17. AccessibleObjectFromWindow Fallback Strategy
- **COM Tunneling**: Never rely solely on `comtypes.client.GetActiveObject` to connect to Office COM. It will frequently throw `MK_E_UNAVAILABLE` errors if Excel is in "Edit Mode" or blocked by Windows DCOM security boundaries. You must always implement a fallback that uses `ctypes.windll.oleacc.AccessibleObjectFromWindow(hwnd, -16)` to tunnel directly into the COM Application object from the active window handle.