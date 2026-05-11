import app from "./app.js";
import { logger } from "./lib/logger.js";
import { mkdirSync } from "fs";

mkdirSync("/tmp/output/items", { recursive: true });
mkdirSync("/tmp/output/analysis", { recursive: true });

const rawPort = process.env["PORT"];
if (!rawPort) throw new Error("PORT environment variable is required");

const port = Number(rawPort);
if (Number.isNaN(port) || port <= 0) throw new Error(`Invalid PORT: "${rawPort}"`);

const hasKey = !!process.env["ANTHROPIC_API_KEY"];

app.listen(port, () => {
  logger.info({ port }, "ArchiveIQ ready");
  logger.info(`API key: ${hasKey ? "YES" : "NO"}`);
});
