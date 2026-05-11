import { Router, type IRouter } from "express";
import healthRouter from "./health";
import lettervaultRouter from "./lettervault";

const router: IRouter = Router();

router.use(healthRouter);
router.use(lettervaultRouter);

export default router;
