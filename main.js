const fs = require('fs');
const got = require('got');
const jsdom = require('jsdom');
const { JSDOM } = jsdom;
const csv = require('csv');
const moment = require('moment');

const url = process.argv.slice(2)[0]
//const url = 'http://artflsrv02.uchicago.edu/cgi-bin/perseus/showrest_?kwic.6.1.5687.0.99.LatinSept18';

got(url).then(response => {
	const dom = new JSDOM(response.body);
	const document = dom.window.document;
	
	const kwicRows = document.querySelectorAll('.content');

	var entries = [];

	for (const kwicRow of kwicRows) {
		const text = kwicRow.textContent
		const title = kwicRow.nextSibling;
		const work = title.querySelector('b');
		const passage = title.querySelector('a');
		
		const entry = 	{	'work': 		work.textContent,
					'passage': 	passage.textContent,
					'text':		text
				};
		// console.log(entry);

		// console.log(work.textContent);
		// const entry = {'citation': title, 
		// console.log(title.textContent);
		// console.log(kwicRow.textContent);

		entries.push(entry);
	}

	//console.log(entries);
	
	serialized = csv.stringify(	entries, 
					{ 	columns: [ { key: 'text' }, {key: 'work'}, {key: 'passage'} ],
						delimiter: '|' 
					});
	serialized.pipe(fs.createWriteStream(`Latin corpus search ${moment().format('MM-DD-YY HH mm ss')}.csv`));


}).catch(err => {
	console.log(err);
});

function kwicText(span) {
	return;
};