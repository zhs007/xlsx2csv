"use strict";

var xlsx = require('node-xlsx');
var fs = require('fs');

var rIntGrammar = new RegExp('[1-9]+[0-9]*');
var rFloatGrammar = new RegExp('[0-9]+[.0-9]*');

function isInt(str) {
    return rIntGrammar.exec(str) == str;
}

function isFloat(str) {
    return rFloatGrammar.exec(str) == str;
}

function arr2phpstr(arr, phpval) {
    let str = '$' + phpval + ' = array(\n';

    for (let ii = 0; ii < arr.length; ++ii) {
        let curs = '';
        for (let key in arr[ii]) {
            if (curs != '') {
                curs += ', ';
            }

            if (arr[ii][key] == undefined) {
                curs += "'" + key + "' => ''";
            }
            else if (typeof arr[ii][key] == 'number') {
                curs += "'" + key + "' => " + arr[ii][key];
            }
            else if (typeof arr[ii][key] == 'string') {
                if (arr[ii][key].indexOf("'") >= 0) {
                    curs += "'" + key + "' => " + '"' + arr[ii][key] + '"';
                }
                else {
                    curs += "'" + key + "' => '" + arr[ii][key] + "'";
                }
            }
            else {
                curs += "'" + key + "' => '" + arr[ii][key] + "'";
            }
        }

        if (ii != 0) {
            str += ',\n\tarray(' + curs + ')';
        }
        else {
            str += '\tarray(' + curs + ')';
        }
    }

    str += '\n);'
    return str;
}

function obj2phpstr(arr, phpval) {
    let str = '$' + phpval + ' = array(\n';
    let ii = 0;

    for (let mkey in arr) {
        let curs = '';
        for (let key in arr[mkey]) {
            if (curs != '') {
                curs += ', ';
            }

            if (arr[mkey][key] == undefined) {
                curs += "'" + key + "' => ''";
            }
            else if (typeof arr[mkey][key] == 'number') {
                curs += "'" + key + "' => " + arr[mkey][key];
            }
            else if (typeof arr[ii][key] == 'string') {
                if (arr[mkey][key].indexOf("'") >= 0) {
                    curs += "'" + key + "' => " + '"' + arr[mkey][key] + '"';
                }
                else {
                    curs += "'" + key + "' => '" + arr[mkey][key] + "'";
                }
            }
            else {
                curs += "'" + key + "' => '" + arr[mkey][key] + "'";
            }
        }

        if (ii != 0) {
            str += ",\n\t'" + mkey + "' => array(" + curs + ")";
        }
        else {
            str += "\t'" + mkey + "' => array(" + curs + ")";
        }

        ++ii;
    }

    str += '\n);'
    return str;
}

function xlsx2php(valname, xlsxfile, jsonfile, excludeline) {
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

        fs.writeFileSync(jsonfile, arr2phpstr(jsonobj, valname), 'utf-8');
    }
}

function xlsx2phpobj(valname, xlsxfile, jsonfile, key) {
    let obj = xlsx.parse(xlsxfile);

    if (obj.length > 0) {
        let jsonobj = {};
        let csvobj = obj[0].data;
        let my = 0;

        for (let y = 0; y < csvobj.length; ++y) {
            if (my == y) {
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

            jsonobj[cobj[key]] = cobj;
        }

        fs.writeFileSync(jsonfile, obj2phpstr(jsonobj, valname), 'utf-8');
    }
}

exports.xlsx2php = xlsx2php;
exports.xlsx2phpobj = xlsx2phpobj;