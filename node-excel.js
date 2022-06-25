
// Requiring the module
const reader = require('xlsx-js-style')
  
// Reading our test file
const file = reader.readFile('./data/list.xlsx')
  
let data = []
  
const sheets = file.SheetNames
  
for(let i = 0; i < sheets.length; i++)
{
   const temp = reader.utils.sheet_to_json(
        file.Sheets[file.SheetNames[i]])
   temp.forEach((res) => {
      data.push(res)
   })
}
  

const file2 = reader.readFile('./data/new.xlsx')

let wscols = [];
let count = 0;
data.map(i => {
	if(count == 0) {
		for (const [key, value] of Object.entries(i)) {
		  if(value.length) {
			wscols = [...wscols, {wch:value.length}]
		  } else {
			wscols = [...wscols, {wch:5}]	  
		  }
		 
		}
		count++;
	} else {
		let u = 0;
		for (const [key, value] of Object.entries(i)) {
		  if(wscols[u].wch < value.length + 5) {
			wscols[u] = {wch:value.length + 5}  
		  }
		  u++
		}
		count++;
	}
})

let num = 2;
let r = 0;
data.map(i => {
	data[r].X1 = {
		t: 'n',
	//	f: '${data[num].X1}'
		f: `=ROUND((1000/((1000/E${num}) + (1000/f${num}))),2)`,
		s: { alignment: { horizontal: "center" } },
		z: "0.00"
	}
	
	data[r].A12 = {
		t: 'n',
	//	f: '${data[num].X1}'
		f: `=ROUND((1000/((1000/E${num}) + (1000/G${num}))),2)`,
		s: { alignment: { horizontal: "center" } },
		z: "0.00"

	}
	
	data[r].X2 = {
		t: 'n',
	//	f: '${data[num].X1}'
		f: `=ROUND((1000/((1000/F${num}) + (1000/G${num}))),2)`,
		s: { alignment: { horizontal: "center" } },
		z: "0.00"

	}
	
	
	
	data[r].P1 = {
		s: { alignment: { horizontal: "center" } },
		f: data[r].P1,
		z: "0.00"
	}
	data[r].P2 = {
		s: { alignment: { horizontal: "center" } },
		f: data[r].P2,
		z: "0.00"
	}
	data[r].X = {
		s: { alignment: { horizontal: "center" } },
		f: data[r].X,
		z: "0.00"
	}
	num++;
	r++;
})

const ws = reader.utils.json_to_sheet(data)
 
let times = data.map(i => i.Time)

// var input = '11:40';
// var parts = input.split(':');
// var minutes = parts[0]*60 +parts[1];
// var inputDate = new Date(minutes * 60 * 1000);
const moment = require('moment');

let date1 = moment("11:40", "hh:mm");
let date  = moment(date1)
let newDate = date1.add(2, 'h');

console.log(date);
console.log(newDate);

// console.log(moment(date - date))
ws['!cols'] = wscols;
reader.utils.book_append_sheet(file,ws,"Sheet2")
reader.writeFile(file,'./data/new.xlsx')
