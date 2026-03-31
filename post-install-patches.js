/**
 * post-install-patches.js
 * ========================
 * Run automatically after `npm install` (via postinstall script in package.json)
 * or manually with `node post-install-patches.js`.
 *
 * Patches node_modules for compatibility with Node 17+ and missing packages
 * that are unavailable in corporate Artifactory registries.
 *
 * Patches applied:
 *   1. node-sass -> dart-sass redirect (avoids native C++ build)
 *   2. sp-build-web Node version check disabled (SPFx 1.12 enforces Node <15)
 *   3. @microsoft/rush-stack-compiler-3.2 shim created (wraps local TypeScript)
 */

var fs = require("fs");
var path = require("path");

var NM = path.join(__dirname, "node_modules");
var patchCount = 0;

function log(msg) {
  console.log("[patch] " + msg);
}

// ---------------------------------------------------------------
// 1. Patch node-sass to delegate to dart-sass (sass package)
//    Also silences the legacy-js-api deprecation warning so the
//    SPFx build system does not treat stderr output as a failure.
// ---------------------------------------------------------------
(function patchNodeSass() {
  var indexPath = path.join(NM, "node-sass", "lib", "index.js");
  if (!fs.existsSync(indexPath)) {
    log("SKIP node-sass patch -- file not found: " + indexPath);
    return;
  }
  var shimContent = [
    '"use strict";',
    "// PATCHED by post-install-patches.js -- delegate to dart-sass",
    'var realSass = require("sass");',
    "var shimModule = {};",
    "shimModule.renderSync = function(options) {",
    "  options.silenceDeprecations = options.silenceDeprecations || [];",
    '  if (options.silenceDeprecations.indexOf("legacy-js-api") === -1) {',
    '    options.silenceDeprecations.push("legacy-js-api");',
    "  }",
    "  return realSass.renderSync(options);",
    "};",
    "shimModule.render = function(options, callback) {",
    "  options.silenceDeprecations = options.silenceDeprecations || [];",
    '  if (options.silenceDeprecations.indexOf("legacy-js-api") === -1) {',
    '    options.silenceDeprecations.push("legacy-js-api");',
    "  }",
    "  return realSass.render(options, callback);",
    "};",
    'shimModule.info = "dart-sass\\t" + (realSass.info || "1.32.13");',
    "module.exports = shimModule;",
    "",
  ].join("\n");
  fs.writeFileSync(indexPath, shimContent, "utf8");
  log("OK  node-sass/lib/index.js -> dart-sass shim (with silenced deprecation)");
  patchCount++;
})();

// ---------------------------------------------------------------
// 2. Patch sp-build-web SPBuildRig.js -- disable Node version check
// ---------------------------------------------------------------
(function patchSPBuildRig() {
  var rigPath = path.join(
    NM,
    "@microsoft",
    "sp-build-web",
    "lib",
    "SPBuildRig.js"
  );
  if (!fs.existsSync(rigPath)) {
    log("SKIP SPBuildRig patch -- file not found: " + rigPath);
    return;
  }
  var content = fs.readFileSync(rigPath, "utf8");

  // Look for the throw statement about Node version
  var throwPattern =
    /throw\s+new\s+Error\s*\(\s*[`'"]Your dev environment is running NodeJS version[^;]*;/;
  if (throwPattern.test(content)) {
    content = content.replace(
      throwPattern,
      'console.warn("WARNING: Running on unsupported Node " + process.version);'
    );
    fs.writeFileSync(rigPath, content, "utf8");
    log("OK  SPBuildRig.js -- Node version check disabled");
    patchCount++;
  } else if (content.indexOf("WARNING: Running on unsupported Node") !== -1) {
    log("SKIP SPBuildRig.js -- already patched");
  } else {
    log("WARN SPBuildRig.js -- could not find the version-check throw statement");
  }
})();

// ---------------------------------------------------------------
// 3. Create @microsoft/rush-stack-compiler-3.2 shim package
// ---------------------------------------------------------------
(function createRushStackCompilerShim() {
  var shimDir = path.join(
    NM,
    "@microsoft",
    "rush-stack-compiler-3.2",
    "lib"
  );
  var pkgPath = path.join(shimDir, "..", "package.json");
  var indexPath = path.join(shimDir, "index.js");

  // Create directory
  if (!fs.existsSync(shimDir)) {
    fs.mkdirSync(shimDir, { recursive: true });
  }

  // package.json
  var pkgJson = JSON.stringify(
    {
      name: "@microsoft/rush-stack-compiler-3.2",
      version: "0.0.0-shim",
      description:
        "Shim package for SPFx 1.12/1.13 builds without the real rush-stack-compiler",
      main: "lib/index.js",
    },
    null,
    2
  );
  fs.writeFileSync(pkgPath, pkgJson + "\n", "utf8");

  // lib/index.js
  var indexContent = [
    '"use strict";',
    "// Shim for @microsoft/rush-stack-compiler-3.2",
    "// Created by post-install-patches.js",
    "",
    'var path = require("path");',
    'var child_process = require("child_process");',
    'var Typescript = require("typescript");',
    "",
    "function TypescriptCompiler(config, buildFolder, terminalProvider) {",
    "  this._config = config || {};",
    "  this._buildFolder = buildFolder;",
    "  this._terminalProvider = terminalProvider;",
    "}",
    "",
    "TypescriptCompiler.prototype.invoke = function () {",
    "  var self = this;",
    "  return new Promise(function (resolve, reject) {",
    '    var tscPath = path.join(path.dirname(require.resolve("typescript")), "..", "bin", "tsc");',
    '    var args = ["--project", path.join(self._buildFolder, "tsconfig.json")];',
    "    if (self._config.customArgs && self._config.customArgs.length) {",
    "      args = args.concat(self._config.customArgs);",
    "    }",
    "    var proc = child_process.spawn(process.execPath, [tscPath].concat(args), {",
    "      cwd: self._buildFolder,",
    '      stdio: ["ignore", "pipe", "pipe"],',
    "      env: process.env",
    "    });",
    '    var stdout = "";',
    '    var stderr = "";',
    '    proc.stdout.on("data", function (data) { stdout += data.toString(); });',
    '    proc.stderr.on("data", function (data) { stderr += data.toString(); });',
    '    proc.on("close", function (code) {',
    '      var output = stdout + "\\n" + stderr;',
    '      var lines = output.split("\\n");',
    "      var errorPattern = /^(.+)\\\\((\\\\d+),(\\\\d+)\\\\):\\\\s+(error|warning)\\\\s+(TS\\\\d+):\\\\s+(.*)$/;",
    "      for (var i = 0; i < lines.length; i++) {",
    "        var match = lines[i].match(errorPattern);",
    "        if (match) {",
    "          var filePath = match[1];",
    "          var line = parseInt(match[2], 10);",
    "          var column = parseInt(match[3], 10);",
    "          var severity = match[4];",
    "          var errorCode = match[5];",
    "          var message = match[6];",
    '          if (severity === "error" && self._config.fileError) {',
    "            self._config.fileError(filePath, line, column, errorCode, message);",
    '          } else if (severity === "warning" && self._config.fileWarning) {',
    "            self._config.fileWarning(filePath, line, column, errorCode, message);",
    "          }",
    "        }",
    "      }",
    "      if (code !== 0) {",
    "        if (stdout.trim()) console.log(stdout);",
    "        if (stderr.trim()) console.error(stderr);",
    '        reject(new Error("TypeScript compilation failed with exit code " + code));',
    "      } else {",
    "        resolve();",
    "      }",
    "    });",
    '    proc.on("error", function (err) { reject(err); });',
    "  });",
    "};",
    "",
    "function TslintRunner(config, buildFolder, terminalProvider) {",
    "  this._config = config || {};",
    "}",
    'TslintRunner.prototype.invoke = function () { return Promise.resolve(); };',
    "",
    "function LintRunner(config, buildFolder, terminalProvider) {",
    "  this._config = config || {};",
    "}",
    'LintRunner.prototype.invoke = function () { return Promise.resolve(); };',
    "",
    "module.exports = {",
    "  Typescript: Typescript,",
    "  TypescriptCompiler: TypescriptCompiler,",
    "  TslintRunner: TslintRunner,",
    "  LintRunner: LintRunner",
    "};",
    "",
  ].join("\n");
  fs.writeFileSync(indexPath, indexContent, "utf8");

  log("OK  @microsoft/rush-stack-compiler-3.2 shim created");
  patchCount++;
})();

// ---------------------------------------------------------------
console.log(
  "\n[patch] Done! Applied " + patchCount + " patch(es) successfully."
);
console.log(
  "[patch] You can now build with:  npm run build"
);
