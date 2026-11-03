// Moodle's mobile-app SSO launch: the URL that hands out a web-service token after the user signs in,
// and the "moodlemobile://token=..." link it answers with. Shared by the CLI and the web page.
//
// The link carries base64("<md5(site + passport)>:::<token>[:::<private token>]").

export interface LaunchToken {
  /** md5(siteUrl + passport) as Moodle computed it; absent when the user pasted a bare token. */
  signature?: string;
  token: string;
  privateToken?: string;
}

/**
 * The launch page for a site. With `confirmed`, Moodle shows a page holding a "Click here if the app
 * does not open automatically" link instead of redirecting straight to the app, so a person in an
 * ordinary browser can copy that link.
 */
export function launchUrl(siteUrl: string, passport: string, opts: { confirmed?: boolean } = {}): string {
  const q = new URLSearchParams({ service: "moodle_mobile_app", passport, urlscheme: "moodlemobile" });
  if (opts.confirmed) q.set("confirmed", "1");
  return `${siteUrl}/admin/tool/mobile/launch.php?${q}`;
}

/** Accepts the copied app link (also URL-encoded or with surrounding text) or a bare 32-hex token. */
export function parseLaunchToken(text: string): LaunchToken {
  const s = decodeURIComponent(text.trim());
  if (/^[a-f0-9]{32}$/i.test(s)) return { token: s };
  const m = /moodlemobile:\/\/token=([A-Za-z0-9+/=_-]+)/.exec(s);
  if (!m) throw new Error('That is not a Moodle app link. It should start with "moodlemobile://token=".');
  let decoded: string;
  try {
    decoded = atob(m[1]!.replace(/-/g, "+").replace(/_/g, "/"));
  } catch {
    throw new Error("The Moodle app link is damaged (not valid base64). Copy it again.");
  }
  const [signature, token, privateToken] = decoded.split(":::");
  if (!token || !/^[A-Za-z0-9]+$/.test(token)) throw new Error("The Moodle app link holds no token. Copy it again.");
  return { signature, token, privateToken: privateToken || undefined };
}
