"use strict";

var xlsx = require('node-xlsx');
var fs = require('fs');

var rIntGrammar = new RegExp('[1-9]+[0-9]*');
var rFloatGrammar = new RegExp('[0-9]+[.0-9]*');

function isInt(str) {
    return rIntGrammar.exec(str);
}

function isFloat(str) {
    return rFloatGrammar.exec(str);
}

function xlsx2json(xlsxfile, jsonfile, excludeline) {
    let obj = xlsx.parse(xlsxfile);

    if (obj.length > 0) {
        let jsonobj = [];
        let csvobj = obj[0].data;
        let my = 0;

        if (excludeline == my) {
            my += 1;
        }

        for (let y = 0; y < csvobj.length; ++y) {
            if (excludeline == y || my == y) {
                continue ;
            }

            let cobj = {};
            for (let x = 0; x < csvobj[y].length; ++x) {
                if (csvobj[y][x] != undefined) {
                    let cstr = csvobj[y][x].toString();
                    if (isInt(cstr)) {
                        cobj[csvobj[my][x]] = parseInt(cstr);
                    }
                    else if (isFloat(cstr)) {
                        cobj[csvobj[my][x]] = parseFloat(cstr);
                    }
                    else {
                        cobj[csvobj[my][x]] = cstr;
                    }
                }
                else {
                    cobj[csvobj[my][x]] = undefined;
                }
            }

            jsonobj.push(cobj);
        }

        fs.writeFileSync(jsonfile, JSON.stringify(jsonobj), 'utf-8');
    }
}

exports.xlsx2json = xlsx2json;