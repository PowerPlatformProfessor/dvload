// Entry point for the Power Platform ToolBox build.
//
// All behaviour lives in the shared task-pane controller; host.ts detects
// the ToolBox bridge globals (window.toolboxAPI / window.dataverseAPI) at
// startup and swaps in the PPTB host + gateway client. This file exists so
// the ToolBox package has its own webpack graph — and so anything genuinely
// PPTB-build-specific has a home that isn't the shared controller.

import "../../addin/src/taskpane/taskpane.js";
