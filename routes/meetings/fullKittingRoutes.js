const express = require("express");
const router = express.Router();
const ctrl = require("../../controllers/meetings/fullKittingController");

router.get("/", ctrl.getFullKittingData);
router.post("/action", ctrl.submitFullKittingAction);

module.exports = router;