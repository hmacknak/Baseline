// Keeps a running add-in on the latest published build.
// The add-in is served from GitHub Pages; every publish writes version.json next to taskpane.html.

export interface LatestVersion {
  version: string;
  label: string;
}

export const CURRENT_VERSION = typeof __BUILD_VERSION__ === "string" ? __BUILD_VERSION__ : "dev";
export const CURRENT_LABEL = typeof __BUILD_LABEL__ === "string" ? __BUILD_LABEL__ : "development";

/** True when a different build is live. Development builds never update themselves. */
export function isUpdate(current: string, latest: string | null | undefined): boolean {
  return !!latest && current !== "dev" && latest !== "dev" && latest !== current;
}

/** Read version.json from the site, bypassing every cache. Resolves null if offline or unreadable. */
export function fetchLatest(): Promise<LatestVersion | null> {
  return new Promise(resolve => {
    try {
      // XMLHttpRequest rather than fetch: older Excel webviews don't have fetch.
      const xhr = new XMLHttpRequest();
      xhr.open("GET", `version.json?t=${Date.now()}`, true);
      xhr.setRequestHeader("Cache-Control", "no-cache");
      xhr.timeout = 8000;
      xhr.onload = () => {
        try {
          const data = JSON.parse(xhr.responseText);
          resolve(xhr.status === 200 && data && typeof data.version === "string" ? data : null);
        } catch {
          resolve(null);
        }
      };
      xhr.onerror = () => resolve(null);
      xhr.ontimeout = () => resolve(null);
      xhr.send();
    } catch {
      resolve(null);
    }
  });
}

/** Load the new build. The query string makes Excel fetch a fresh taskpane.html instead of a cached one. */
export function applyUpdate(version: string): void {
  window.location.replace(`${window.location.pathname}?v=${encodeURIComponent(version)}`);
}

const AUTO_KEY = "baseline.autoUpdatedTo";

/**
 * On opening: if a newer build is live, switch to it once. Guarded per version so a stale CDN copy
 * can never cause a reload loop.
 */
export async function autoUpdateOnOpen(): Promise<boolean> {
  if (CURRENT_VERSION === "dev") return false;
  const latest = await fetchLatest();
  if (!latest || !isUpdate(CURRENT_VERSION, latest.version)) return false;
  try {
    if (sessionStorage.getItem(AUTO_KEY) === latest.version) return false;
    sessionStorage.setItem(AUTO_KEY, latest.version);
  } catch {
    // No session storage: don't risk a loop; the banner and button still offer the update.
    return false;
  }
  applyUpdate(latest.version);
  return true;
}
