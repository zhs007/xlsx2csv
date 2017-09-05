#!/usr/bin/env node

"use strict";

var fs = require('fs');
var process = require('process');
var glob = require('glob');
var xlsx2json = require('../xlsx2json');
var argv = require('yargs')
    .option('t', {
        alias : 'tablename',
        demand: false,
        describe: 'use tablename',
        type: 'boolean'
    })
    .option('e', {
        alias : 'exclude',
        demand: false,
        describe: 'exclude same line',
        type: 'int'
    })
    .option('k', {
        alias : 'key',
        demand: false,
        describe: 'build obj with key',
        type: 'string'
    })
    .usage('Usage: xlsx2json path')
    .example('xlsx2json path', 'xlsx2json path')
    .help('h')
    .alias('h', 'help')
    .epilog('copyright 2015')
    .argv;

var basearr = argv._;

if (basearr == undefined || basearr.length < 1) {
    console.log('Usage: xlsx2json path');

    process.exit(1);
}

var excludeline = -1;
if (argv.hasOwnProperty('exclude')) {
    excludeline = argv.exclude;
}

var mkey = undefined;
if (argv.hasOwnProperty('key')) {
    mkey = argv.key;
}

for (var j = 0; j < basearr.length; ++j) {
    var lstfile = glob.sync(basearr[j]);
    for (var i = 0; i < lstfile.length; ++i) {
        var srcfile = lstfile[i];
        if (fs.existsSync(srcfile)) {
            if (srcfile.slice(srcfile.length - 4) == '.xls' || srcfile.slice(srcfile.length - 5) == '.xlsx') {
                var filename = srcfile;
                var ptindex = srcfile.lastIndexOf('.');
                if (ptindex > 0) {
                    filename = srcfile.slice(0, ptindex);
                }

                if (mkey == undefined) {
                    xlsx2json.xlsx2json(srcfile, filename + '.json', excludeline);
                }
                else {
                    xlsx2json.xlsx2jsonobj(srcfile, filename + '.json', mkey);
                }

                console.log(srcfile + ' OK!');
            }
        }
    }
}

process.exit();