'use strict';

require('dotenv').config();
var fs = require('fs');
var path = require('path');
var build = require('@microsoft/sp-build-web');

build.addSuppression("Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be made available via the local class names object passed to the 'require' call of the battery. In order to refer to this class in your module, you must use the global CSS class name directly.");

/* ---------- Use dart-sass instead of node-sass ---------- */
var sassTask = build.sass;
if (sassTask && sassTask.taskConfig) {
  sassTask.taskConfig.useCSSModules = true;
}

/* ---------- Inject SHAREPOINT_SITE_URL from .env into serve.json ---------- */
var siteUrl = process.env.SHAREPOINT_SITE_URL;
if (siteUrl) {
  var serveConfigPath = path.join(__dirname, 'config', 'serve.json');
  var serveConfig = JSON.parse(fs.readFileSync(serveConfigPath, 'utf8'));
  serveConfig.initialPage = siteUrl.replace(/\/+$/, '') + '/_layouts/workbench.aspx';
  fs.writeFileSync(serveConfigPath, JSON.stringify(serveConfig, null, 2) + '\n', 'utf8');
}

build.initialize(require('gulp'));
