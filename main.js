const fs = require('fs');
const got = require('got');
const jsdom = require('jsdom');
const { JSDOM } = jsdom;
const excelJS = require('excelJS');
const moment = require('moment');

const url = process.argv.slice(2)[0]

got(url).then(response => {
	const dom = new JSDOM(response.body);
	const document = dom.window.document;
	
	const kwicRows = document.querySelectorAll('.content');

	const workbook = new excelJS.Workbook();
	const worksheet = workbook.addWorksheet('Data');
	worksheet.columns = [
		{ header: 'Text', key: 'text', width: 75, style: { alignment: { wrapText: true }, font: { name: 'Times New Roman' }, border: { right: { style: 'thin' } } } },
		{ header: 'Work', key: 'work', width: 15 },
		{ header: 'Passage', key: 'passage', width: 15}
	];
	worksheet.views = [ { zoomScale: 180 } ]

 	for (const kwicRow of kwicRows) {
		const text = kwicRow.textContent
		const title = kwicRow.nextSibling;
		const work = title.querySelector('b');
		const passage = title.querySelector('a');

		worksheet.addRow({	text: text.trim(),
					work: work.textContent.trim(),
					passage: passage.textContent.trim()
				});
	}

	worksheet.getRow(1).font = { bold: true };
	worksheet.getRow(1).border = { bottom: { style: 'thin' } };
	
	filename = `Latin corpus search ${moment().format('MM-DD-YY HH mm ss')}.xlsx`;

	workbook.xlsx.writeFile(filename).then(() => {
		console.log(filename);
	}).catch((err) => {
		console.log(err);
	});
	// serialized.pipe(fs.createWriteStream(filename, {mode: '0444'}));

}).catch(err => {
	console.log(err);
});
