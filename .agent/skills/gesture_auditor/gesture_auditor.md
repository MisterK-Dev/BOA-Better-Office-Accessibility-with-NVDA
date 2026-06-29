# Agent Skill: Keystroke & Gesture Auditing Pipeline

**Purpose:**
This skill acts as a keyboard sanity check for the BOA add-on. It ensures that any new keyboard shortcut is conflict-free, follows prefix mode guidelines, is properly exposed to the NVDA Input Gestures dialog for user customization, and uses the correct unified category `BOA (Better Office Accessibility)`.

---

## Keyboard Design Rules for BOA

### 1. The Single Global Default Key
* **Rule:** The only key BOA may register globally in NVDA's key map by default is the Prefix trigger key: `NVDA+e` (or whatever the user changes it to).
* **Rule:** No other default direct keys (e.g., binding a feature directly to `Ctrl+Shift+O` in the code) should be shipped. Let the user decide if they want a direct hotkey.

### 2. Standalone Script Exposure
* **Rule:** Every user-facing, keyboard-driven feature (like sheet reordering, cell monitoring, layout analysis) **must** be exposed as a standalone `script_` method on its respective `AppModule` class.
* **Rule:** Every exposed script must have a descriptive, translatable docstring (which NVDA shows as the script's description in the Input Gestures dialog).
* **Rule:** Every exposed script must be assigned to the unified category: `_("BOA (Better Office Accessibility)")` (e.g. `scriptCategory = _("BOA (Better Office Accessibility)")`).

### 3. Exclude Prefix Helpers
* **Rule:** Internal prefix plumbing methods (like key interceptors and escape cancelers) must **not** have docstrings. This keeps them hidden from the Input Gestures dialog so users don't bind hotkeys to them by mistake.
* **Exempt List**:
  * `script_handleCommandKey`
  * `script_cancelCommandPrefix`
  * `script_evaluateAndRead`
  * `script_caret_nextParagraph`
  * `script_caret_previousParagraph`

---

## How to Execute the Gesture Audit

To statically analyze the codebase for keyboard mapping issues, run the following command from the root of the repository:

```powershell
python .agent/skills/gesture_auditor/scripts/audit_gestures.py
```

### Interpretation of Results

The auditing script will output a detailed report:
1. **Registered Scripts**: A list of all scripts found grouped by category.
2. **Exposure Check**: Verification that all prefix commands in the manager files have corresponding scripts in the `AppModule` class.
3. **Conflicts**: Detailed list of duplicate keys or invalid categories.

If conflicts or missing exposures are found, the script will exit with code `1`. The audit must return `0` conflicts before finalizing any release.
