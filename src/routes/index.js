fileUpload = require("../controllers/filesController");

exports.appRoute = router => {
  router.get("/files", fileUpload.showUploadPage);
  router.post("/upload", fileUpload.uploadExcel);
};
