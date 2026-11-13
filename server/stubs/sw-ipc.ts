// Desktop-only assistant IPC is not part of the engine service: inference runs on
// the platform, and the browser talks to it directly. The retained main modules
// (Sheets and Slides) import these at load time, so they stay no-ops here.
export function registerSheetsControlledAiIpc(): void {
  /* intentionally empty */
}
export function registerCompanionAiIpc(): void {
  /* intentionally empty */
}
