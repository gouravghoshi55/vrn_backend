const express = require("express");
const router = express.Router();
const upload = require("../../middleware/upload");
const ctrl = require("../../controllers/meetings/agreementController");

router.get("/", ctrl.getAgreementData);

router.post("/action", upload.single("uploadPdf"), ctrl.submitAgreementAction);

module.exports = router;