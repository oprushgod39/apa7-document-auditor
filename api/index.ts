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
 * Note: the API kicks off the processing pipeline (`processSession`) and
 * returns 202 immediately, then the client polls `/status`. That background
 * work is wrapped in `waitUntil` (see server/src/run_background.ts) so this
 * function stays alive until the pipeline settles, bounded by the
 * `maxDuration` set for this function in vercel.json.
 */
export default function handler(req: IncomingMessage, res: ServerResponse): void {
  app(req as never, res as never);
}
