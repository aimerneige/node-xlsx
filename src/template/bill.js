import xlsx from "xlsx-js-style";

const wb = xlsx.utils.book_new();

const data = [
  ["科目总账"],
  ["2023年"],
  ["2023年", "", "摘要", "借方", "贷方", "方向", "余额"],
  ["月", "日"],
];

const ws = xlsx.utils.aoa_to_sheet(data);

const cellStyles = {
  font: { bold: true }, // 设置字体加粗
  alignment: {
    horizontal: "center", // 水平居中
    vertical: "center", // 垂直居中
  },
};

let titleCellAddresses = ["A1", "A3", "A4", "B4", "C3", "D3", "E3", "F3", "G3"];

titleCellAddresses.forEach((cellAddress) => {
  ws[cellAddress].s = cellStyles;
});

ws["!merges"] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 6 } }, // 科目总账
  { s: { r: 2, c: 0 }, e: { r: 2, c: 1 } }, // 2023年
  { s: { r: 2, c: 2 }, e: { r: 3, c: 2 } }, // 摘要
  { s: { r: 2, c: 3 }, e: { r: 3, c: 3 } }, // 借方
  { s: { r: 2, c: 4 }, e: { r: 3, c: 4 } }, // 贷方
  { s: { r: 2, c: 5 }, e: { r: 3, c: 5 } }, // 方向
  { s: { r: 2, c: 6 }, e: { r: 3, c: 6 } }, // 余额
];

ws["!cols"] = [
  { wch: 8 },
  { wch: 8 },
  { wch: 30 },
  { wch: 12 },
  { wch: 12 },
  { wch: 8 },
  { wch: 12 },
];

ws["!rows"] = [{ hpx: 24 }];

xlsx.utils.book_append_sheet(wb, ws, "科目总账");

const filePath = "./output/bill.xlsx";
xlsx.writeFile(wb, filePath);
