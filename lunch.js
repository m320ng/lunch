var exec = require('child_process').exec;
var execSync = require('exec-sync');
var async = require('async');
var excelParser = require('excel-parser');
var moment = require('moment');

var xlsfile = __dirname+'/lunch.xls';
var ment1file = __dirname+'/ment1.mp3';
var ment2file = __dirname+'/ment2.mp3';

var tasks = [];

tasks.push(function(next) {
	execSync('rm -f '+xlsfile);	
	execSync('rm -f '+ment1file);	
	execSync('rm -f '+ment2file);	
	exec('wget -q -U Mozila -O '+xlsfile+' http://hs.eipark.co.kr/main/food_menu_excel.aspx', function(err, stdout, stderr) {
		if (err) {
			console.log(err);
			return;
		}
		next();
	});
});

var sheetId;
var now = {month:moment().month()+1, date:moment().date()};
var klunch = [];
var flunch = [];

tasks.push(function(next) {
	excelParser.worksheets({
		inFile: xlsfile
	}, function(err, worksheets){
		if(err) {
			console.error(err);
			return;
		}
		worksheets.forEach(function(sheet) {
			if (sheet.name=='식단표') {
				sheetId = sheet.id;
				return;
			}
		});
		next();
	});
});

tasks.push(function(next) {
	if (!sheetId) {
		return;
	}
	excelParser.parse({
		inFile: xlsfile,
		worksheet: sheetId,
		skipEmpty: false
	},function(err, records){
		if(err) {
			console.error(err);
			return;
		}
		parseLunch(records);
		//console.log(records);
		next();
	});
});

async.series(tasks, function() {
	var ment = [];
	var ment2 = [];
	ment.push(now.month+'월 '+now.date+'일 점심메뉴는, ');
	if (klunch.join('')) {
		ment.push('한식은 '+klunch.join(', ')+' 입니다.');
	} else {
		ment.push('한식은 없습니다.');
	}
	if (flunch.join('')) {
		ment2.push('퓨전은 '+flunch.join(', ')+' 입니다.');
	} else {
		ment2.push('퓨전은 없습니다.');
	}
	//console.log(klunch);
	//console.log(flunch);
	console.log(ment.join(' '));
	console.log(ment2.join(' '));
	var ret1 = execSync('wget -q -U Mozilla -O '+ment1file+' "http://translate.google.com/translate_tts?ie=UTF-8&tl=ko&q='+ment.join(' ')+'"');
	if (ret1) {
		console.log(ret1);
	}
	var ret2 = execSync('wget -q -U Mozilla -O '+ment2file+' "http://translate.google.com/translate_tts?ie=UTF-8&tl=ko&q='+ment2.join(' ')+'"');

	if (ret2) {
		console.log(ret2);
	}
	exec('mpg321 -q '+ment1file+' '+ment2file, function(err, stdout, stderr) {
		console.log(err);
		console.log('complete');
	});
});

function parseLunch(records) {
	var step = '';
	var month;
	var date;
	var selectIndex;
	records.forEach(function(items, index) {
		if (!items.length) {
			console.log(items);
			return;
		}
		if (!step) {
			items.forEach(function(item, itemidx) {
				var matches = item.match(/([0-9]+)월[ ]*([0-9]+)일/);
				if (matches) {
					var m = parseInt(matches[1]);
					var d = parseInt(matches[2]);
					step = 'date';
					month = '';
					date = '';
					if (m==now.month && d==now.date) {
						month = m;
						date = d;
						selectIndex = itemidx;
					}
				}
			});
		} else if (step=='date') {
			items.forEach(function(item, itemidx) {
				if (/한[ ]*식/.test(item)) {
					klunch.push(items[selectIndex]);
					step = 'klunch';
					return;
				}
			});
		} else if (step=='klunch') {
			items.forEach(function loop(item, itemidx) {
				if (loop.stop) return;
				if (/퓨[ ]*전/.test(item)) {
					flunch.push(items[selectIndex]);
					step = 'flunch';
					loop.stop = true;
					return;
				} else {
					if (item && itemidx==selectIndex) {
					klunch.push(item);
					}
				}
			});
		} else if (step=='flunch') {	
			items.forEach(function loop(item, itemidx) {
				if (loop.stop) return;
				if (/김[ ]*치/.test(item)) {
					step = 'end';
					loop.stop = true;
					return;
				} else {
					if (item && itemidx==selectIndex) {
					flunch.push(item);
					}
				}
			});
		}
	});
}
