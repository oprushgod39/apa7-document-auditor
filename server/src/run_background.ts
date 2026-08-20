import { waitUntil } from "@vercel/functions";

/**
 * Runs `promise` in the background without blocking the response.
 *
 * On Vercel, a serverless function's execution (including any detached
 * async work started during the request) ends once the response is sent —
 * a plain fire-and-forget promise is not guaranteed to complete. `waitUntil`
 * keeps the function alive until `promise` settles (bounded by the
 * function's `maxDuration`), so the polling contract (`202` now, `/status`
 * polled later) keeps working.
 *
 * Outside Vercel (local dev, Docker, tests) there is no serverless request
 * context to extend, so this stays a plain detached promise — the same
 * behavior as before, appropriate for a long-running process.
 */
export function runInBackground(promise: Promise<unknown>): void {
  const settled = promise.catch(() => {
    /* errors are captured on the session by the caller */
  });
  if (process.env.VERCEL) {
    waitUntil(settled);
  }
}
