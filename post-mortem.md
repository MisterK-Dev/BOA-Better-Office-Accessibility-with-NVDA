# Architectural Choices & Preferences
- **Cell Monitor (v1.4 dev 1):** We chose to use `comtypes.client.GetEvents` to hook into Excel's native `SheetChange` and `SheetCalculate` events instead of running a background polling timer. This ensures zero CPU overhead when Excel is idle, instantly catching any value change regardless of the trigger.
- **COM Thread Safety Constraints:** Due to the single-threaded nature of NVDA, we strictly route all COM event callbacks through `queueHandler.queueFunction()` before interacting with the NVDA `ui.message` pipeline to avoid core UI thread locks or deadlocks when Excel fires an asynchronous change event.

# Active Mistakes & Bug Mitigations
- None recorded yet. If the `comtypes` Event Hook proves unstable during Cell Editing Mode, we may have to pivot to a polling timer, but the initial design favors native event hooks for performance.
