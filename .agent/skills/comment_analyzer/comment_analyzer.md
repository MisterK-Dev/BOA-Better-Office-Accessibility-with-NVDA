# Comment Analyzer Developer Skill

This developer skill defines the automated workflow for auditing, tracking, and correcting source code comments, docstrings, and documentation flags in the codebase.

## 1. Core Principles
* **Manual Execution**: This skill runs only when explicitly invoked by the user.
* **Audit-First Mode**: The skill must always start in a read-only audit mode and generate a report before making any changes.
* **Local Scope Only (No Git Commit/Push)**: All changes, corrections, and file edits made during the correction phase must remain strictly local and unstaged. The agent MUST NOT run any `git add`, `git commit`, or `git push` commands. The user must manually review and commit the changes.
* **Functional Integrity**: Adding or correcting comments must never break the code's execution, syntax, or functionality.

---

## 2. Phase 1: Read-Only Audit
Scan the requested directory or files to evaluate the status of comments and documentation. Compile all findings into a report file named `comment_audit_report.md` in the artifacts directory.

### Audits to Perform:
1. **Commented-Out Code (Dead Code)**:
   * Scan for active lines of code that are commented out (e.g., lines containing disabled function calls, loops, imports, or variable assignments). 
   * Flag these as code clutter that should be cleaned up.
2. **Missing Docstrings**:
   * Flag all public classes, methods, and functions that lack docstrings.
   * Verify existing docstrings: Check if all parameters and return types defined in the function signature are present and correctly documented in the docstring. Flag any mismatches.
3. **Missing Explanatory Comments**:
   * Identify logical blocks with high complexity that lack explanatory comments. This includes:
     * Complex nested conditional branches or loops.
     * Direct Windows API calls (DLL calls via `ctypes` or `windll`).
     * Direct COM object bindings and dispatches.
     * Custom asynchronous thread operations (e.g. `wx.CallAfter` or deferred execution).
4. **Localization & Translation Safety**:
   * Inspect all string literals wrapped in the gettext translation function `_()`.
   * Flag any translation calls that contain developer comments, HTML tags, or formatting code, as these pollute translation catalogs.
5. **AGGREGATION: TODO & Task Flags**:
   * Search for flags such as `TODO`, `FIXME`, `HACK`, `XXX`, or `BUG`.
   * Gather them into a list mapped by file path and line number.

---

## 3. Phase 2: Interactive Correction Mode
Once the report `comment_audit_report.md` is presented to the user, pause execution and wait for explicit instructions.

### Guidelines for Correcting Comments:
* **Interactive Approval**: Only fix missing comments or docstrings that the user has explicitly requested to be fixed.
* **Grounding & Accuracy**:
   * Analyze the actual logic of the function, the parameter names, and return values.
   * Generate highly accurate docstrings. Never hallucinate or write generic/redundant comments (e.g. avoid comments like "sets value" for a setter method). Explain the *intent* of the code.
   * If a complex block is ambiguous, do not guess. Flag it and ask the user for clarification.
* **Formatting Alignment**:
   * Match the file's existing indentation system (e.g., 4 spaces or tabs) exactly to prevent syntax/indentation errors.
   * Use standard PEP 8 spacing guidelines for inline comments (e.g., two spaces after code, followed by `#` and a single space).

---

## 4. Verification & Safety Protocols
After applying any corrections, the agent must run the following verification steps:

1. **Python Syntax Check**:
   * Run Python's compiler module in check mode on all modified files to ensure syntax is 100% valid:
     `python -m py_compile <filename>`
2. **SCons Compilation Check**:
   * Run the add-on compilation tool to ensure the package builds without errors:
     `scons`
   * If compilation fails, revert the changes immediately.
3. **Git Diff Review**:
   * Run a diff to review the exact changes:
     `git diff`
   * Confirm that only comment lines were added or modified, and no functional code logic, statements, or variable names were altered.
4. **Stop for Verification**:
   * Notify the user that the corrections have been applied locally and present the updated files for verification. DO NOT commit or push.
