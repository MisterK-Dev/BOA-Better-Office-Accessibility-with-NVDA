# Agent Skill: Security Auditor

**Trigger Phase:**
Automatically triggered during the Release Validation pipeline (Step 6), or when specifically requested by the user.

**Purpose:**
This skill performs a strictly read-only security audit of the entire NVDA Add-on codebase. It uses Python's Abstract Syntax Tree (`ast`) to intelligently scan for common Python vulnerabilities and NVDA-specific risks (e.g., COM freezing loops, eval injection) without triggering false positives on code comments.

---

## Execution Steps

### 1. Execute the Auditor Script
- Run the python script located at `.agent/skills/security_auditor/scripts/audit_security.py`.
- **Command:** `python .agent/skills/security_auditor/scripts/audit_security.py`
- This script scans all `.py` files inside the `addon/` directory.

### 2. Report Aggregation
- **Do not halt the pipeline** if the script outputs warnings.
- The script is read-only and will never modify the codebase.
- Collect all output warnings (file names, line numbers, and issue descriptions).
- Append the formatted list of vulnerabilities to the final "Greenlight" report presented to the Developer during release validation.

### 3. Await Developer Instruction
- You are strictly forbidden from writing fixes for detected vulnerabilities unless explicitly instructed by the Developer. Wait for the Developer to review the report and provide manual permission to address specific warnings.
