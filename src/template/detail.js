import xlsx from "xlsx-js-style";

const wb = xlsx.utils.book_new();

const data = [
  ["科目明细账"],
  ["11901 粮食"],
  ["2023年01月-12月", "科目： 11901 粮食"],
  ["日期", "", "", "凭证号", "摘要", "借方", "贷方", "方向", "余额"],
  ["年", "月", "日"],
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
  "A4",
  "D4",
  "E4",
  "F4",
  "G4",
  "H4",
  "I4",
  "A5",
  "B5",
  "C5",
];

titleCellAddresses.forEach((cellAddress) => {
  ws[cellAddress].s = cellStyles;
});

ws["!merges"] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } }, // 科目明细账
  { s: { r: 3, c: 0 }, e: { r: 3, c: 2 } }, // 日期
  { s: { r: 3, c: 3 }, e: { r: 4, c: 3 } }, // 凭证号
  { s: { r: 3, c: 4 }, e: { r: 4, c: 4 } }, // 摘要
  { s: { r: 3, c: 5 }, e: { r: 4, c: 5 } }, // 借方
  { s: { r: 3, c: 6 }, e: { r: 4, c: 6 } }, // 贷方
  { s: { r: 3, c: 7 }, e: { r: 4, c: 7 } }, // 方向
  { s: { r: 3, c: 8 }, e: { r: 4, c: 8 } }, // 余额
];

ws["!cols"] = [
  { wch: 5 },
  { wch: 5 },
  { wch: 5 },
  { wch: 10 },
  { wch: 20 },
  { wch: 10 },
  { wch: 10 },
  { wch: 5 },
  { wch: 10 },
];

ws["!rows"] = [{ hpx: 24 }];

xlsx.utils.book_append_sheet(wb, ws, "明细账");

const filePath = "./output/detail.xlsx";
xlsx.writeFile(wb, filePath);
