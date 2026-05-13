const express = require("express");
const router = express.Router();
const ctrl = require("../../controllers/meetings/meetingsSubController");

router.get("/", ctrl.getMeetingsSub);
router.post("/action", ctrl.submitMeetingsSubAction);

module.exports = router;