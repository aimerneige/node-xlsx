import xlsx from "xlsx-js-style";

const wb = xlsx.utils.book_new();

const data = [
  ["总账余额表"],
  ["2023年01月-12月", "科目： 所有科目"],
  ["科目代码", "科目名称", "期初", "", "本期", "", "期末", "", "累计", ""],
  ["", "", "借方", "贷方", "借方", "贷方", "借方", "贷方", "借方", "贷方"],
];

const ws = xlsx.utils.aoa_to_sheet(data);

const cellStyles = {
  font: { bold: true }, // 设置字体加粗
  alignment: {
    horizontal: "center", // 水平居中
    vertical: "center", // 垂直居中
  },
};

let titleCellAddresses = [
  "A1",
  "A3",
  "B3",
  "C3",
  "E3",
  "G3",
  "I3",
  "C4",
  "D4",
  "E4",
  "F4",
  "G4",
  "H4",
  "I4",
  "J4",
];

titleCellAddresses.forEach((cellAddress) => {
  ws[cellAddress].s = cellStyles;
});

ws["!merges"] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }, // 总账余额表
  { s: { r: 2, c: 0 }, e: { r: 3, c: 0 } }, // 科目代码
  { s: { r: 2, c: 1 }, e: { r: 3, c: 1 } }, // 科目名称
  { s: { r: 2, c: 2 }, e: { r: 2, c: 3 } }, // 期初
  { s: { r: 2, c: 4 }, e: { r: 2, c: 5 } }, // 本期
  { s: { r: 2, c: 6 }, e: { r: 2, c: 7 } }, // 期末
  { s: { r: 2, c: 8 }, e: { r: 2, c: 9 } }, // 累计
];

ws["!cols"] = [{ wch: 14 }, { wch: 24 }];

ws["!rows"] = [{ hpx: 24 }];

xlsx.utils.book_append_sheet(wb, ws, "总账余额表");

const filePath = "./output/total.xlsx";
xlsx.writeFile(wb, filePath);
