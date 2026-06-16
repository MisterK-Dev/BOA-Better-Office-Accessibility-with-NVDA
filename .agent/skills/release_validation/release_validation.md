# Agent Skill: Release Validation Pipeline

**Trigger Phase:**
Automatically trigger this skill whenever the user asks to push a commit to the `main` branch, finalize a stable build, or execute a version release.

**Purpose:**
This skill acts as an unskippable gatekeeper. It forces the AI Agent to pause, run a strict pre-release validation checklist, and explicitly request human "Greenlight" approval before executing any `git push` commands. This guarantees absolute compliance with NVDA Add-on standards and deployment safety protocols.

---

## The Pre-Release Checklist

Before communicating with the user, the Agent MUST autonomously execute and verify the following 5 steps:

### 1. Translation Wrapping Verification (`_()`)
- Scan all newly added or modified `.py` files.
- **Rule:** Every single user-facing English string (UI labels, dialog text, speech announcements) MUST be wrapped in the GNU `gettext` function `_()`.
- **Rule:** No raw f-strings may be used for user text. They must use `.format()`.
- **Action:** If any raw English UI strings are found without `_()`, halt validation, report the exact file and line number to the user, and ask to fix it.

### 2. GPL-2.0 Header Audit
- Check the top 4 lines of every newly added or modified `.py` file.
- **Rule:** They must contain the strict GPL-2.0 Copyright header defined in `rules.md`.
- **Action:** If missing, automatically append the header and note it in the final summary.

### 3. Version Synchronization
- Open `buildVars.py`, `CHANGELOG.md`, and `addon/doc/en/readme.md`.
- **Rule:** The version string (e.g., `v1.6.0`) must be identical across all three files.
- **Action:** If there is a mismatch (e.g., `buildVars.py` is `1.6.0` but `readme.md` still says `1.5.0`), halt validation and alert the user.

### 4. The Scons Mock Build
- Execute the terminal command: `scons`
- **Rule:** The build must complete successfully, generating the `.nvda-addon` file without any Python tracebacks or `msgfmt` compilation errors.
- **Action:** If it fails, trigger the `Build Fixer` skill. Do not proceed until `scons` exits cleanly.

### 5. Keystroke & Gesture Audit
- **Rule:** Run the [Gesture Auditing Skill](file:///.agent/skills/gesture_auditor/gesture_auditor.md) to ensure all user gestures are properly categorized under `BOA (Better Office Accessibility)` and there are zero shortcut conflicts.
- **Action:** Open and follow the instructions in `gesture_auditor.md` to execute the audit. The audit must pass with 0 conflicts and verify that all keyboard commands are correctly exposed before releasing.

---

## The "Greenlight" Protocol (Mandatory Gate)

Once the checklist is 100% complete and passing, the Agent MUST NOT push the code yet. 

The Agent must output a final message to the user containing:
1. A summary of the 5 checklist items (e.g., *"GPL Headers: Verified", "Scons Build: Passed", "Gesture Audit: Passed"*).
2. The exact Git commands it intends to run (e.g., `git push origin main`, `git push origin v1.7.0`).
3. The exact Question: **"Everything is verified. Do I have your final Greenlight to push this release to the remote repository?"**

**Wait for Explicit Approval:**
The Agent is strictly forbidden from executing `git push` until the user replies with an explicit "Yes", "Go ahead", or "Greenlight".
