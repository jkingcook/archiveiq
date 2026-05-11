import { Router } from "express";
import { getBusSummary } from "../lib/intelligence-bus.js";

const router = Router();

router.get("/", (_req, res) => {
  res.json(getBusSummary());
});

export default router;
