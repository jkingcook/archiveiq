import { Router } from "express";
import { BUS, createProject } from "../lib/intelligence-bus.js";

const router = Router();

router.get("/", (_req, res) => {
  const projects = Object.values(BUS.projects).map(p => ({
    ...p,
    item_count: BUS.itemStore.filter(i => i.project_id === p.id).length,
    row_count: BUS.itemStore
      .filter(i => i.project_id === p.id)
      .reduce((s, i) => s + ((i as { register_rows?: unknown[] }).register_rows?.length ?? 0), 0),
    is_active: BUS.activeProject === p.id,
  }));
  res.json({ projects, activeProject: BUS.activeProject });
});

router.post("/create", (req, res) => {
  const { name, type, description, primary_language, date_range_from, date_range_to, primary_researcher, institution, notes } = req.body;
  if (!name) return res.status(400).json({ error: "name is required" });

  const project = createProject({
    name,
    type: type ?? "Custom",
    description: description ?? "",
    primary_language: primary_language ?? "English",
    date_range_from: date_range_from ?? "",
    date_range_to: date_range_to ?? "",
    primary_researcher: primary_researcher ?? "",
    institution: institution ?? "",
    notes: notes ?? "",
  });

  res.json({ success: true, project });
});

router.post("/set-active", (req, res) => {
  const { projectId } = req.body;
  if (!projectId || !BUS.projects[projectId]) {
    return res.status(404).json({ error: "Project not found" });
  }
  BUS.activeProject = projectId;
  BUS.projects[projectId].last_active = new Date().toISOString();
  BUS.sessionLog.push({ timestamp: new Date().toISOString(), action: "set_active", project_id: projectId });
  res.json({ success: true, activeProject: projectId });
});

router.put("/:id", (req, res) => {
  const { id } = req.params;
  if (!BUS.projects[id]) return res.status(404).json({ error: "Not found" });
  BUS.projects[id] = { ...BUS.projects[id], ...req.body, id };
  res.json({ success: true, project: BUS.projects[id] });
});

router.delete("/:id", (req, res) => {
  const { id } = req.params;
  if (id === "machen-family-papers") {
    return res.status(403).json({ error: "Cannot delete the Machen Family Papers project" });
  }
  if (!BUS.projects[id]) return res.status(404).json({ error: "Not found" });
  delete BUS.projects[id];
  if (BUS.activeProject === id) {
    BUS.activeProject = Object.keys(BUS.projects)[0] ?? null;
  }
  res.json({ success: true });
});

export default router;
