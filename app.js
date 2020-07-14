//Node.js ==> Express Framework (SIMPLE SERVER)
require('core-js/modules/es.promise');
require('core-js/modules/es.string.includes');
require('core-js/modules/es.object.assign');
require('core-js/modules/es.object.keys');
require('core-js/modules/es.symbol');
require('core-js/modules/es.symbol.async-iterator');
require('regenerator-runtime/runtime');
const express = require("express");
let app = express();
//Port to listen on
const PORT = 3000;

const path = require("path");

const bodyParser = require("body-parser");
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs/dist/es5');
const bootstrap = require("./src/boostrap");

// default options
app.use(fileUpload());

//Use a Custom Templating Engine
app.set("view engine", "pug");

app.set("views", path.resolve("./src/views"));

//Request Parsing
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
//Create Express Router
const router = express.Router();
app.use(router);

const rootPath = path.resolve("./dist");
app.use(express.static(rootPath));

bootstrap(app, router);

//Main Page (Home)
router.get("/", (req, res, next) => {
  return res.send("Hello There");
});

router.use((err, req, res, next) => {
  if (err) {
    //Handle file type and max size of image
    return res.send(err.message);
  }
});

app.listen(PORT, err => {
  if (err) return console.log(`Cannot Listen on PORT: ${PORT}`);
  console.log(`Server is Listening on: http://localhost:${PORT}/`);
});
