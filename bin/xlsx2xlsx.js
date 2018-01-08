#!/usr/bin/env node

"use strict";

var fs = require('fs');
var process = require('process');
var glob = require('glob');
var xlsx2xlsx = require('../xlsx2xlsx');
var argv = require('yargs')
    .option('t', {
        alias : 'tablename',
        demand: false,
        describe: 'use tablename',
        type: 'boolean'
    })
    .option('t', {
        alias : 'tablename2',
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
    .usage('Usage: xlsx2xlsx path')
    .example('xlsx2xlsx path', 'xlsx2xlsx path')
    .help('h')
    .alias('h', 'help')
    .epilog('copyright 2015')
    .argv;

var basearr = argv._;

if (basearr == undefined || basearr.length < 1) {
    console.log('Usage: xlsx2xlsx path');

    process.exit(1);
}

var excludeline = -1;
if (argv.hasOwnProperty('exclude')) {
    excludeline = argv.exclude;
}

if(basearr.length > 0)
{
    var name1 = glob.sync(basearr[0])[0];
    var name2 = glob.sync(basearr[1])[0];
    var param1;
    var param2;
    if (fs.existsSync(name1)) {
        if (name1.slice(name1.length - 4) == '.xls' || name1.slice(name1.length - 5) == '.xlsx') {
            var ptindex = name1.lastIndexOf('.');
            if (ptindex > 0) {
                param1 = name1.slice(0, ptindex);
            }

        }
    }

    xlsx2xlsx.xlsx2xlsx(name1, name2, param1 + "_g.xlsx", excludeline);
}

process.exit();