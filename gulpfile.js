"use strict";

/* eslint global-require: "off" */

const gulp = require('gulp');
const eslint = require("gulp-eslint");
const jsdoc2md = require("jsdoc-to-markdown");
const toc = require('markdown-toc');
const fs = require("fs");
const Jasmine = require("jasmine");
const { series, parallel } = gulp;


const PATHS = {
    lib: "./lib/**/*.js",
    unit: "./test/unit/**/*.js",
    karma: ["./test/helpers/**/*.js", "./test/unit/**/*.spec.js"], // Helpers need to go first
    examples: "./examples/**/*.js",
    browserify: {
        source: "./lib/XlsxPopulate.js",
        base: "./browser",
        bundle: "xlsx-populate.js",
        noEncryptionBundle: "xlsx-populate-no-encryption.js",
        sourceMap: "./",
        encryptionIngores: ["./lib/Encryptor.js"]
    },
    readme: {
        template: "./docs/template.md",
        build: "./README.md"
    },
    blank: {
        workbook: "./blank/blank.xlsx",
        template: "./blank/template.js",
        build: "./lib/blank.js"
    },
    jasmineConfigs: {
        unit: "./test/unit/jasmine.json",
        e2eGenerate: "./test/e2e-generate/jasmine.json",
        e2eParse: "./test/e2e-parse/jasmine.json"
    }
};

PATHS.lint = [PATHS.lib];
PATHS.unitTestSources = [PATHS.lib, PATHS.unit];

// Function to clear the require cache as running unit tests mess up later tests.
const clearRequireCache = () => {
    for (const moduleId in require.cache) {
        delete require.cache[moduleId];
    }
};

const runJasmine = (configPath, cb) => {
    process.chdir(__dirname);
    clearRequireCache();
    const jasmine = new Jasmine();
    jasmine.loadConfigFile(configPath);
    jasmine.onComplete(passed => cb(null));
    jasmine.execute();
};

gulp.task("blank", () => {
    return Promise
        .all([
            fs.readFileAsync(PATHS.blank.workbook, "base64"),
            fs.readFileAsync(PATHS.blank.template, "utf8")
        ])
        .spread((data, template) => {
            const output = template.replace("{{DATA}}", data);
            return fs.writeFileAsync(PATHS.blank.build, output);
        });
});

gulp.task("lint", () => {
    return gulp.src(PATHS.lint)
        .pipe(eslint())
        .pipe(eslint.format());
});

gulp.task("unit", cb => {
    runJasmine(PATHS.jasmineConfigs.unit, cb);
});

gulp.task("e2e-generate", cb => {
    runJasmine(PATHS.jasmineConfigs.e2eGenerate, cb);
});

gulp.task("e2e-parse", cb => {
    runJasmine(PATHS.jasmineConfigs.e2eParse, cb);
});


gulp.task("docs", cb => {
    // eslint-disable-next-line no-sync
    let text = fs.readFileSync(PATHS.readme.template, "utf8");
    const tocText = toc(text, { filter: str => str.indexOf('NOTOC-') === -1 }).content;
    text = text.replace("<!-- toc -->", tocText);
    text = text.replace(/NOTOC-/mg, "");
    // eslint-disable-next-line no-sync
    fs.writeFileSync(PATHS.readme.build, text);
    cb();
});
