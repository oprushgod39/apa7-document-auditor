import express from "express";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import path from "node:path";
import fs from "node:fs";
import { config } from "./config.js";
import { apiRouter, errorHandler } from "./api/routes.js";

export function createApp(): express.Express {
  const app = express();
  app.disable("x-powered-by");

  app.use(
    helmet({
      contentSecurityPolicy: {
        directives: {
          defaultSrc: ["'self'"],
          scriptSrc: ["'self'"],
          styleSrc: ["'self'", "'unsafe-inline'"],
          imgSrc: ["'self'", "data:"],
          connectSrc: ["'self'"],
          objectSrc: ["'none'"],
          frameAncestors: ["'none'"],
        },
      },
    })
  );

  app.use(
    "/api",
    rateLimit({
      windowMs: config.rateLimitWindowMs,
      max: config.rateLimitMax,
      standardHeaders: true,
      legacyHeaders: false,
      message: {
        error: {
          code: "RATE_LIMITED",
          message: "Too many requests. Please wait a moment and try again.",
        },
      },
    })
  );

  app.use(express.json({ limit: "256kb" }));
  app.use("/api", apiRouter());

  // Serve the built frontend when available (production / docker).
  if (fs.existsSync(config.frontendDistDir)) {
    app.use(express.static(config.frontendDistDir));
    app.get(/^(?!\/api).*/, (_req, res) => {
      res.sendFile(path.join(config.frontendDistDir, "index.html"));
    });
  }

  app.use(errorHandler);
  return app;
}
