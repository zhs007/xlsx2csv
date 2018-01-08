"use strict";

var xlsx = require('node-xlsx');
var fs = require('fs');

var lstSymbols = [];
var lstNumber = [];

function change(str) {
	var str2;
    var index = -1;
	for(var i = 0; i < lstSymbols.length; ++i)
	{
		if(lstSymbols[i] == str)
		{
			index = i;
			break;
		}
	}


	str2 = lstNumber[index];

	return str2;
}

function xlsx2xlsx(t1file, p1file, t2file, excludeline) {
    var t1obj = xlsx.parse(t1file);
    var p1obj = xlsx.parse(p1file);

	if(p1obj.length > 0) {

        var p1obj_data = p1obj[0].data;

        for (var x = 1; x < p1obj_data.length; ++x) {
            if (p1obj_data[x][0] != undefined) {
                lstNumber[x - 1] = p1obj_data[x][0];
                lstSymbols[x - 1] = p1obj_data[x][1];
            }
        }


        var t1obj_data = t1obj[0].data;

        var reels = [];
        var reels2 = [];

        for (var y = 0; y < t1obj_data.length; ++y) {

            for (var x = 0; x < t1obj_data[y].length; ++x) {
                if (t1obj_data[y][x] != undefined) {

                    if (reels[y] == undefined) {
                        reels[y] = [];

                        reels2[y] = [];
                    }

                    var str = t1obj_data[y][x].toString();

                    reels[y][x] = str;

                    var str2 = change(str);
                    reels2[y][x] = str2;
                }
            }
        }


        for (var i = 0; i < reels2.length; ++i) {
            for (var j = 0; j < reels2[i].length; ++j) {
                if(reels2[i][j] == undefined) {
                    reels2[i][j] = -1;
                }
            }
        }

        for (var i = 0; i < reels2.length; ++i) {
        	var le = reels2[i].length;

            for (var j = 0; j < reels[i].length; ++j) {
                reels2[i][le + j] = reels[i][j];
            }
        }


        console.log("over");

        for(var i = 0; i < reels2.length; ++i)
		{
			reels2[i].splice(0, 0, i);
		}

		var lstline = ["line"];
        for(var i = 0; i < reels[0].length; ++i)
		{
			lstline[lstline.length] = "R" + (i + 1);
		}
        for(var i = 0; i < reels[0].length; ++i)
        {
            lstline[lstline.length] = "S" + (i + 1);
        }
        reels2.splice(0, 0, lstline);


        var buffer = xlsx.build([{name: "Sheet1", data: reels2}]); // Returns a buffer

        fs.writeFileSync(t2file, buffer);
    }

}

exports.xlsx2xlsx = xlsx2xlsx;