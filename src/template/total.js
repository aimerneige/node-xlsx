import xlsx from "xlsx";

const data = [
  ["总账余额表"],
  ["2023年01月-12月", "科目： 所有科目"],
  ["科目代码", "科目名称", "期初", "本期", "期末", "累计"],
  ["101", "现金"],
];

const ws = xlsx.utils.aoa_to_sheet(data);

ws["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }];

const wb = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(wb, ws, "Sheet 1");

const filePath = "./output/total.xlsx";
xlsx.writeFile(wb, filePath);
