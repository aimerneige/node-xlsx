import xlsx from 'node-xlsx';
import * as fs from 'fs';

const data = [
  [1, 2, 3],
  [true, false, null, 'sheetjs'],
  ['foo', 'bar', new Date('2014-02-19T14:30Z'), '0.3'],
  ['baz', null, 'qux'],
];
var buffer = xlsx.build([{name: 'mySheetName', data: data}]);

const filePath = "./output/demo_example.xlsx";
fs.writeFile(filePath, buffer, function(err) {
  if(err) {
      return console.log(err);
  }
  console.log("The file was saved!");
});