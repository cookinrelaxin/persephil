const fs = require('fs');
const got = require('got');
const jsdom = require('jsdom');
const { JSDOM } = jsdom;
const excelJS = require('excelJS');
const moment = require('moment');

const url = process.argv.slice(2)[0]
const assert = require('assert');

got(url).then(response => {
	const dom = new JSDOM(response.body);
	const document = dom.window.document;
	
	const kwicRows = document.querySelectorAll('.content');

	const workbook = new excelJS.Workbook();
	const worksheet = workbook.addWorksheet('Data');
	worksheet.columns = [
		{ 	header: 'Text', 
			key: 'text', 
			width: 75, 
			style: { 
				alignment: { wrapText: true }, 
				border: { right: { style: 'thin' } } 
			} 
		},
		{ 	header: 'Work', 
			key: 'work', 
			width: 15 
		},
		{ 	header: 'Passage', 
			key: 'passage', 
			width: 15
		}
	];
	worksheet.views = [ { zoomScale: 180 } ]

 	for (const kwicRow of kwicRows) {
		const title = kwicRow.nextSibling;
		const work = title.querySelector('b');
		const passage = title.querySelector('a');

		let stk = [];
		stk.push(kwicRow);
		const richText = [];

		while(stk.length !== 0){

			let currentNode = stk.pop();

			if (currentNode && currentNode.childNodes && currentNode.childNodes.length >0){

				for(let i = currentNode.childNodes.length - 1; i>=0; i--){
					stk.push(currentNode.childNodes[i]);
				}
			}
			else {
				if (currentNode.parentNode.parentNode.tagName == 'B') {
					richText.push({font: { bold: true, name: 'Times New Roman' }, text: currentNode.textContent });
				}
				else {
					richText.push({font: { bold: false, name: 'Times New Roman' }, text: currentNode.textContent });
				}
			}
		}
		richText[0].text = richText[0].text.trimStart();
		if (richText[0].text === '') { richText.shift(); }
		
		richText[richText.length-1].text = richText[richText.length-1].text.trimEnd();
		if (richText[richText.length-1].text === '') { richText.pop(); }

		worksheet.addRow({	text: {richText: richText},
		 			work: work.textContent.trim(),
		 			passage: passage.textContent.trim()
				});
	}

	worksheet.getRow(1).font = { bold: true };
	worksheet.getRow(1).border = { bottom: { style: 'thin' } };

	const cell = worksheet.getCell('A2');
	assert(cell.type == excelJS.ValueType.RichText);

	
	filename = `Latin corpus search ${moment().format('MM-DD-YY HH mm ss')}.xlsx`;

	workbook.xlsx.writeFile(filename).then(() => {
		console.log(filename);
	}).catch((err) => {
		console.log(err);
	});

}).catch(err => {
	console.log(err);
});
