import express, { type Request, type Response, type NextFunction } from "express";
import { join, dirname } from "path";
import { fileURLToPath } from "url";
import pinoHttp from "pino-http";
import { logger } from "./lib/logger.js";
import projectsRouter from "./routes/projects.js";
import setupRouter from "./routes/setup.js";
import processRouter from "./routes/process.js";
import analyzeRouter from "./routes/analyze.js";
import intelligenceRouter from "./routes/intelligence.js";
import exportRouter from "./routes/export.js";
import statusRouter from "./routes/status.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const isProduction = process.env.NODE_ENV === "production";

const BASE = (process.env["BASE_PATH"] ?? "/archiveiq").replace(/\/$/, "");

app.use(pinoHttp({ logger }));
app.use(express.json({ limit: "50mb" }));

const publicDir = isProduction
  ? join(__dirname, "public")
  : join(__dirname, "../src/public");

app.use(BASE, express.static(publicDir));

const apiBase = `${BASE}/api`;

app.use(apiBase + "/projects", projectsRouter);
app.use(apiBase + "/setup", setupRouter);
app.use(apiBase + "/process", processRouter);
app.use(apiBase + "/items", processRouter);
app.use(apiBase + "/analyze", analyzeRouter);
app.use(apiBase + "/analysis", analyzeRouter);
app.use(apiBase + "/intelligence", intelligenceRouter);
app.use(apiBase + "/export", exportRouter);
app.use(apiBase + "/status", statusRouter);

app.get(`${BASE}/{*path}`, (_req, res) => {
  res.sendFile(join(publicDir, "index.html"));
});

app.get("/", (_req, res) => {
  res.redirect(BASE);
});

// Global JSON error handler — catches unhandled Express/multer errors
// Must be defined after all routes
app.use((err: Error & { status?: number }, _req: Request, res: Response, _next: NextFunction) => {
  logger.error({ err: err.message, stack: err.stack }, "Unhandled Express error");
  res.status(err.status ?? 500).json({ error: err.message ?? "Internal server error" });
});

export default app;
