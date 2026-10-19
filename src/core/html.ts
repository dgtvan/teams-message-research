/**
 * Shared HTML-to-text decoding for message content — used for both the
 * real-time push payload (trouter.ts) and DOM-extracted content, which
 * both carry Teams' raw HTML.
 *
 * This module used to also hold a REST-based catch-up path
 * (fetchMessagesSince) that called Teams' internal chatsvc API directly.
 * Confirmed live that it doesn't work: the endpoint 401s on every call —
 * it needs an explicit Skype-protocol auth token that lives encrypted in
 * localStorage and isn't attached by page.request or an in-page fetch,
 * and real Teams navigation no longer calls this endpoint itself either,
 * so there's no live request to harvest a working header from. It failed
 * silently (empty result, no error) regardless of how long the gap was.
 * Catch-up now reuses the same DOM-based open-and-extract path as the
 * backup visit loop instead (see catchUpAllTracked in connectors/teams-web/index.ts).
 */
export function htmlToText(html: string): string {
  return (
    html
      // Turn block-level tag boundaries into a line break BEFORE stripping
      // tags — otherwise two adjacent block-level elements (e.g. Teams'
      // quoted/reply-preview block immediately followed by the actual
      // message text) collapse into one run-on string with zero characters
      // between them once the tags are gone. Confirmed via a captured
      // message where the quoted block and the real reply text were fused
      // mid-word ("rồivụ") — the raw HTML from the Trouter push payload is
      // regex-stripped here with nothing preserving where one element
      // ended and the next began. Not an exhaustive list of block tags,
      // just the ones Teams' payloads are known/likely to use — a missed
      // tag just costs a lost line break, the same failure mode this had
      // before, not a new one.
      .replace(/<\s*(div|p|br|li|tr)[^>]*>/gi, "\n")
      .replace(/<\/\s*(div|p|li|tr)\s*>/gi, "\n")
      .replace(/<[^>]*>/g, "")
      .replace(/&nbsp;/g, " ")
      .replace(/&amp;/g, "&")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&quot;/g, '"')
      .replace(/&#39;/g, "'")
      // Collapse horizontal whitespace per line (not newlines — those are
      // the block-boundary separators just inserted above), then dedupe
      // consecutive blank lines down to a single one so a run of empty
      // <div></div> separators doesn't turn into a wall of blank lines.
      .split("\n")
      .map((line) => line.replace(/\s+/g, " ").trim())
      .filter((line, i, arr) => line !== "" || (i > 0 && arr[i - 1] !== ""))
      .join("\n")
      .trim()
  );
}
