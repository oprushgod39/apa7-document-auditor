import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import puppeteer, { type Browser } from "puppeteer-core";
import { AppError } from "../errors.js";

/**
 * True when running inside a serverless function (Vercel or plain AWS
 * Lambda, which Vercel's Node runtime is built on) rather than a developer's
 * own machine. Vercel sets VERCEL=1 for both build and runtime; the Lambda
 * env vars are a fallback for other serverless hosts on the same runtime.
 */
function isServerlessRuntime(): boolean {
  return Boolean(process.env.VERCEL || process.env.AWS_LAMBDA_FUNCTION_NAME || process.env.LAMBDA_TASK_ROOT);
}

/** Common install locations for a Chromium-family browser, by platform. */
function candidateLocalExecutables(): string[] {
  const home = os.homedir();
  if (process.platform === "win32") {
    const programFiles = process.env["PROGRAMFILES"] ?? "C:\\Program Files";
    const programFilesX86 = process.env["PROGRAMFILES(X86)"] ?? "C:\\Program Files (x86)";
    const localAppData = process.env["LOCALAPPDATA"] ?? path.join(home, "AppData", "Local");
    return [
      path.join(programFiles, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(programFilesX86, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(localAppData, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(programFilesX86, "Microsoft", "Edge", "Application", "msedge.exe"),
      path.join(programFiles, "Microsoft", "Edge", "Application", "msedge.exe"),
    ];
  }
  if (process.platform === "darwin") {
    return [
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      "/Applications/Chromium.app/Contents/MacOS/Chromium",
      "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    ];
  }
  return [
    "/usr/bin/google-chrome-stable",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium-browser",
    "/usr/bin/chromium",
    "/usr/bin/microsoft-edge",
    "/snap/bin/chromium",
  ];
}

function findLocalExecutable(): string | null {
  for (const candidate of candidateLocalExecutables()) {
    try {
      if (fs.existsSync(candidate)) return candidate;
    } catch {
      /* ignore */
    }
  }
  return null;
}

interface LaunchPlan {
  executablePath: string;
  args: string[];
}

/**
 * Version-pinned remote Chromium pack for @sparticuz/chromium-min.
 *
 * `@sparticuz/chromium` (the full package) bundles its brotli-compressed
 * Chromium + shared-library archives under node_modules/@sparticuz/chromium/bin
 * and reads them via fs at runtime rather than a static import. Vercel's
 * build-time file tracer doesn't follow that dynamic fs read reliably, even
 * with an explicit `includeFiles` glob in vercel.json — the deployed
 * function still launched Chromium without its bundled libnss3.so and
 * friends. `chromium-min` sidesteps Vercel's bundler entirely: it fetches
 * this prebuilt pack over HTTPS on cold start and extracts it to /tmp
 * itself, independent of what Vercel decided to include in the function.
 * Keep this URL's version suffix in lockstep with the installed
 * @sparticuz/chromium-min version.
 */
const CHROMIUM_PACK_URL =
  "https://github.com/Sparticuz/chromium/releases/download/v131.0.1/chromium-v131.0.1-pack.tar";

async function resolveLaunchPlan(): Promise<LaunchPlan> {
  if (isServerlessRuntime()) {
    const chromium = (await import("@sparticuz/chromium-min")).default;
    const executablePath = await chromium.executablePath(CHROMIUM_PACK_URL);
    return { executablePath, args: chromium.args };
  }
  const executablePath = process.env.PUPPETEER_EXECUTABLE_PATH || findLocalExecutable();
  if (!executablePath) {
    throw new AppError(
      "PROCESSING_FAILED",
      "No local Chrome or Edge installation was found to render the merged PDF. Install Google Chrome, or set PUPPETEER_EXECUTABLE_PATH to a Chromium-based browser's executable.",
      500
    );
  }
  return { executablePath, args: ["--no-sandbox", "--disable-setuid-sandbox"] };
}

/** Renders a self-contained HTML document (no external resources) to PDF. */
export async function renderHtmlToPdf(html: string): Promise<Buffer> {
  const plan = await resolveLaunchPlan();
  let browser: Browser;
  try {
    browser = await puppeteer.launch({
      executablePath: plan.executablePath,
      args: plan.args,
      headless: true,
    });
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new AppError("PROCESSING_FAILED", `The PDF renderer could not start (${message}).`, 500);
  }
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    const pdf = await page.pdf({
      format: "letter",
      printBackground: true,
      margin: { top: "1in", bottom: "1in", left: "1in", right: "1in" },
    });
    return Buffer.from(pdf);
  } finally {
    await browser.close().catch(() => {});
  }
}
