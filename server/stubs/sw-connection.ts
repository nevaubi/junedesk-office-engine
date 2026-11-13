// The engine service never calls a model. Any retained code path that reaches
// for the desktop Bedrock connection fails loudly instead of contacting AWS.
const unavailable = (): never => {
  throw new Error("Model inference is not available in the Office engine service.");
};
export const CONFIG_EXAMPLE = {};
export function connectionPath(): string {
  return "";
}
export function loadConnection(): never {
  return unavailable();
}
export function configuredProfile(): never {
  return unavailable();
}
export async function streamWriter(): Promise<never> {
  return unavailable();
}
export async function searchPublic(): Promise<never> {
  return unavailable();
}
