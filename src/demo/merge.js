const XLSX = require('xlsx');

// 创建一个新的工作表
const ws_data = [
  ["Header1", "Header2", "Header3"],
  ["Data1", "Data2", "Data3"],
  ["Data4", "Data5", "Data6"]
];

// 创建一个工作簿
const ws = XLSX.utils.aoa_to_sheet(ws_data);

// 合并单元格
// 在此示例中，我们合并 A1 到 B1 的单元格
ws['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }  // 合并 A1 和 B1
];

// 创建一个新的工作簿
const wb = XLSX.utils.book_new();

// 将工作表添加到工作簿中
XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

// 导出为 Excel 文件
XLSX.writeFile(wb, "output/demo_merge.xlsx");

console.log('Excel 文件已生成');
