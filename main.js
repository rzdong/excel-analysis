let xlsx = require('node-xlsx');
let fs = require('fs');
let path = require('path');

const assets = 'file/';

const files = fs.readdirSync(path.join(__dirname, assets))
files.forEach((item, index) => {
	let sheets = xlsx.parse(assets + item);


	for (let k = 0; k < sheets.length; k++) {

		/**
		 * sheet: {
		 * 		name: '',
		 * 		data: [
		 * 			[     ], 每行数据
		 * 			[     ], 每行数据
		 * 		]
		 * }
		 */
		for (let i = 0; i < sheets[k]['data'].length; i++) {

			for (let j = 0; j < sheets[k]['data'][i].length; j++) {
				let unit = sheets[k]['data'][i][j];

				if (unit === 'NaN') {
					try {
						if (sheets[k]['data'][i - 1][j] !== 'NaN' && sheets[k]['data'][i + 1] && sheets[k]['data'][i + 1][j] !== 'NaN') {
							sheets[k]['data'][i][j] = Number(((Number(sheets[k]['data'][i - 1][j]) + Number(sheets[k]['data'][i + 1][j])) / 2).toFixed(2));
						} else if (sheets[k]['data'][i - 1][j] !== 'NaN' && sheets[k]['data'][i + 2] && sheets[k]['data'][i + 2][j] !== 'NaN') {

							let delta = Number(((Number(sheets[k]['data'][i + 2][j]) - Number(sheets[k]['data'][i - 1][j])) / 3).toFixed(2));
							sheets[k]['data'][i][j] = Number((Number(sheets[k]['data'][i - 1][j]) + delta).toFixed(2));
							sheets[k]['data'][i + 1][j] = Number((Number(sheets[k]['data'][i - 1][j]) + delta * 2).toFixed(2));
							// console.log(typeof delta, sheets[k]['data'][i][j], sheets[k]['data'][i + 1][j])
						} else if (sheets[k]['data'][i - 1][j] !== 'NaN' && sheets[k]['data'][i + 3] && sheets[k]['data'][i + 3][j] !== 'NaN') {
							let delta = Number(((Number(sheets[k]['data'][i + 3][j]) - Number(sheets[k]['data'][i - 1][j])) / 4).toFixed(2));
							sheets[k]['data'][i][j] = Number((Number(sheets[k]['data'][i - 1][j]) + delta).toFixed(2));
							sheets[k]['data'][i + 1][j] = Number((Number(sheets[k]['data'][i - 1][j]) + delta * 2).toFixed(2));
							sheets[k]['data'][i + 2][j] = Number((Number(sheets[k]['data'][i - 1][j]) + delta * 3).toFixed(2));
						} else {

						}
					} catch (error) {
						console.log(error);
						console.log(`出错了！错误所在表${sheets[k]}表, 计算第${i+1}行, 第${j+1}列。`);
					}


				}

			}

		}
	}
	var buffer = xlsx.build(sheets);




	fs.writeFile('target/' + item, buffer, function (err) {
		if (err) {
			console.log("Write failed: " + err);
			return;
		}
		console.log("写入" + item + "完成");
	});
})

