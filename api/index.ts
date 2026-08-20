import type { IncomingMessage, ServerResponse } from "node:http";
// Import the server's TypeScript source directly (not server/dist) so this
// file type-checks against real types and Vercel's esbuild-based function
// bundler can trace + compile it without a separate pre-build step.
import { createApp } from "../server/src/app.js";

const app = createApp();

/**
 * Vercel Node serverless entrypoint. An Express app is directly callable as
 * `(req, res) => void`, matching the Node runtime's request-handler
 * convention, so we hand it the request/response objects unchanged.
 *
 * vercel.json rewrites every `/api/*` request to this function; Express's
 * own router (mounted at `/api` in `createApp`) does the rest of the
 * routing internally.
 *
 * Note: this wraps a single long-running pipeline (`processSession`) that
 * the API kicks off and returns 202 from immediately, then polls via
 * `/status`. On Vercel, a serverless function's execution — including any
 * detached async work started during the request — ends when the response
 * is sent (or at `maxDuration`), so long documents must finish processing
 * before the handler returns for this fire-and-forget pattern to complete
 * reliably. This entrypoint does not change that behavior; it only adapts
 * the existing Express app to run as a function.
 */
export default function handler(req: IncomingMessage, res: ServerResponse): void {
  app(req as never, res as never);
}
