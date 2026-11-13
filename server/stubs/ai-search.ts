// Genspark cloud search is not used by the firm platform. Retained code that
// imports it gets inert values; nothing here contacts the network.
export const gskApiKey = (): string => "";
export function setGskProxyUrl(): void {
  /* no proxy */
}
export async function gskSlideGenerate(): Promise<never> {
  throw new Error("Cloud generation is not available in the Office engine service.");
}
export async function gskWebSearch(): Promise<never> {
  throw new Error("Cloud search is not available in the Office engine service.");
}
export async function gskImageSearch(): Promise<never> {
  throw new Error("Cloud search is not available in the Office engine service.");
}
export const GSK_AVAILABLE = false;
