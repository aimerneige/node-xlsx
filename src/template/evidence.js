import xlsx from "xlsx-js-style";

const wb = xlsx.utils.book_new();

const data = [
  ["记 帐 凭 证"],
  [],
  ["", "   20xx-x-x"],
  ["核算单位:[001]                 有限公司", "", "", "第     号"],
  ["摘     要", "会 计 科 目", "借方金额", "贷方金额"],
  ["文本", "示例", 4124.241, 41.4],
  ["文本", "示例", 1451461.241, 1.4],
  ["文本", "示例", 15145.241, -21351.4],
  ["文本", "示例", 13141.241, -135.4],
  ["文本", "示例", 145161.241, 1516.4],
  ["附单据数              张", "合计:肆仟肆佰伍拾元整", "", ""],
  [
    "财务主管:                 记帐:             复核:             出纳:             制单:     经办人:  ",
    "",
    "",
    "",
  ],
];

const ws = xlsx.utils.aoa_to_sheet(data);

ws["C11"] = { f: "=SUM(C6:C10)" };
ws["D11"] = { f: "=SUM(D6:D10)" };

let centerStyle = {
  font: {
    bold: true, // 设置字体加粗
    name: "楷体_GB2312", // 楷体
    sz: 11, // 大小
  },
  alignment: {
    horizontal: "center", // 水平居中
    vertical: "center", // 垂直居中
  },
};

["B3", "D4", "A5", "B5", "C5", "D5", "C11", "D11"].forEach((a) => {
  ws[a].s = centerStyle;
});

["A1"].forEach((a) => {
  let titleStyle = JSON.parse(JSON.stringify(centerStyle));
  titleStyle.font.sz = 28;
  titleStyle.alignment.vertical = "bottom";
  ws[a].s = titleStyle;
});

["A5", "B5", "C5", "D5"].forEach((a) => {
  let subTitleStyle = JSON.parse(JSON.stringify(centerStyle));
  subTitleStyle.font.sz = 14;
  ws[a].s = subTitleStyle;
});

["B3", "D4"].forEach((a) => {
  let biggerFontsStyle = JSON.parse(JSON.stringify(centerStyle));
  biggerFontsStyle.font.sz = 12;
  ws[a].s = biggerFontsStyle;
});

["C6", "D6", "C7", "D7", "C8", "D8", "C9", "D9", "C10", "D10"].forEach((a) => {
  let numberStyle = JSON.parse(JSON.stringify(centerStyle));
  numberStyle.alignment.vertical = "bottom";
  numberStyle.alignment.horizontal = "end";
  numberStyle.numFmt = "#,##0.00";
  ws[a].s = numberStyle;
});

["C11", "D11"].forEach((a) => {
  let sumStyle = JSON.parse(JSON.stringify(centerStyle));
  sumStyle.alignment.horizontal = "end";
  sumStyle.numFmt = "#,##0.00";
  ws[a].s = sumStyle;
});

["A4", "A11", "B11"].forEach((a) => {
  let leftStyle = JSON.parse(JSON.stringify(centerStyle));
  leftStyle.alignment.horizontal = "start";
  ws[a].s = leftStyle;
});

["A6", "B6", "A7", "B7", "A8", "B8", "A9", "B9", "A10", "B10", "A12"].forEach(
  (a) => {
    let leftBottomStyle = JSON.parse(JSON.stringify(centerStyle));
    leftBottomStyle.alignment.horizontal = "start";
    leftBottomStyle.alignment.vertical = "bottom";
    ws[a].s = leftBottomStyle;
  }
);

ws["!merges"] = [
  { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }, // 记账凭证
  { s: { r: 3, c: 0 }, e: { r: 3, c: 1 } }, // 核算单位
  { s: { r: 11, c: 0 }, e: { r: 11, c: 3 } }, // 财务主管
];

ws["!cols"] = [{ wch: 32 }, { wch: 32 }, { wch: 20 }, { wch: 20 }];

ws["!rows"] = [
  { hpx: 36 },
  { hpx: 3 },
  { hpx: 30.75 },
  { hpx: 27.75 },
  { hpx: 33 },
  { hpx: 33 },
  { hpx: 33 },
  { hpx: 33 },
  { hpx: 33 },
  { hpx: 33 },
  { hpx: 33 },
  { hpx: 33 },
];

xlsx.utils.book_append_sheet(wb, ws, "记账凭证");

const filePath = "./output/evidence.xlsx";
xlsx.writeFile(wb, filePath);
