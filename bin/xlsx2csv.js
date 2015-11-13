#!/usr/bin/env node

var fs = require('fs');
var process = require('process');
var glob = require('glob');
var xlsx2csv = require('../xlsx2csv');
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
    .usage('Usage: xlsx2csv path')
    .example('xlsx2csv path', 'xlsx2csv path')
    .help('h')
    .alias('h', 'help')
    .epilog('copyright 2015')
    .argv;

var basearr = argv._;

if (basearr == undefined || basearr.length != 1) {
    console.log('Usage: xlsx2csv path');

    process.exit(1);
}

var lstfile = glob.sync(basearr[0]);
for (var i = 0; i < lstfile.length; ++i) {
    var srcfile = lstfile[i];
    if (fs.existsSync(srcfile)) {
        if (srcfile.slice(srcfile.length - 4) == '.xls' || srcfile.slice(srcfile.length - 5) == '.xlsx') {
            var filename = srcfile;
            var ptindex = srcfile.lastIndexOf('.');
            if (ptindex > 0) {
                filename = srcfile.slice(0, ptindex);
            }

            xlsx2csv.xlsx2csv(srcfile, filename + '.csv');

            console.log(srcfile + ' OK!');
        }
    }
}