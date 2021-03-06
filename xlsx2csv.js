"use strict";

var xlsx = require('node-xlsx');
var fs = require('fs');

function xlsx2csv(xlsxfile, csvfile, excludeline) {
    var obj = xlsx.parse(xlsxfile);

    if (obj.length > 0) {
        var csvf = fs.openSync(csvfile, 'w');

        var csvobj = obj[0].data;

        for (var y = 0; y < csvobj.length; ++y) {
            if (excludeline == y) {
                continue ;
            }

            for (var x = 0; x < csvobj[y].length; ++x) {
                if (csvobj[y][x] != undefined) {
                    var str = csvobj[y][x].toString();

                    if (str.indexOf(',') >= 0) {
                        fs.writeSync(csvf, '"');
                        fs.writeSync(csvf, csvobj[y][x]);
                        fs.writeSync(csvf, '"');
                    }
                    else {
                        fs.writeSync(csvf, csvobj[y][x]);
                    }
                }

                if (x < csvobj[y].length - 1) {
                    fs.writeSync(csvf, ',');
                }
            }
            fs.writeSync(csvf, '\r\n');
        }

        fs.closeSync(csvf);
    }
}

exports.xlsx2csv = xlsx2csv;