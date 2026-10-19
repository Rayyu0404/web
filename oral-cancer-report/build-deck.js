const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require("/root/.claude/skills/synced/e531a0c6-b3ad-4d78-b93d-73691016111d_850ad4c0-f2ce-497e-95a0-f1c1f3109215/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "oral-cancer.pptx";
const FONT = "Microsoft JhengHei";

const HEX = {
  ink: "2A1520", berry: "7A1F3D", rose: "F6E6EB", white: "FFFFFF",
  teal: "0F766E", grey: "6B5B62", mint: "E3F2EF", gold: "C2410C",
};
const THEME = {
  name: "Oral Cancer Report",
  headFontFace: FONT,
  bodyFontFace: FONT,
  colors: {
    dk1: HEX.ink, lt1: HEX.white, dk2: HEX.berry, lt2: HEX.rose,
    accent1: HEX.berry, accent2: HEX.teal, accent3: HEX.gold, accent4: "B98A9A",
    accent5: "4B5563", accent6: "9CA3AF", hlink: HEX.teal, folHlink: HEX.berry,
  },
};

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.title = "口腔癌 分組報告";
pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
const C = pres.SchemeColor;

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "DARK",
  background: { color: HEX.ink },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.5, w: 8.8, h: 1.2, fontSize: 40, bold: true, color: C.background1, valign: "middle", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 2.75, w: 8.8, h: 1.2, fontSize: 18, color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.white },
  margin: [0.5, 0.5, 0.5, 0.5],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9.0, h: 0.75, fontSize: 28, bold: true, color: C.text2, valign: "middle", margin: 0 }, text: "" } },
    { text: { text: "口腔癌 分組報告", options: { x: 0.5, y: 5.25, w: 4, h: 0.25, fontSize: 9, color: "8C7A82", margin: 0 } } },
  ],
  slideNumber: { x: 9.0, y: 5.25, w: 0.5, h: 0.25, fontSize: 9, color: "8C7A82", align: "right" },
});

function contentSlide(title, section) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  return s;
}
function txt(s, text, opts) {
  s.addText(text, Object.assign({ isTextBox: true, fontFace: FONT, margin: 0 }, opts));
}
function card(s, x, y, w, h, fill, name) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08, objectName: name });
}
async function iconCircle(s, Comp, x, y, d, bg, fg) {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { color: bg } });
  const pad = d * 0.25;
  s.addImage({ data: await icon(Comp, fg), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad });
}
const SRC = (s, t) => txt(s, "資料來源：" + t, { x: 0.5, y: 4.92, w: 9.0, h: 0.28, fontSize: 9, color: HEX.grey });

(async () => {
  // 1 Title
  pres.addSection({ title: "開場" });
  let s = pres.addSlide({ masterName: "DARK", sectionTitle: "開場" });
  s.addText("口腔癌", { placeholder: "title" });
  s.addText("認識台灣男性第 4 大癌症：成因、症狀、預防與治療", { placeholder: "body" });
  await iconCircle(s, fa.FaTooth, 7.9, 0.5, 1.2, HEX.berry, HEX.white);
  txt(s, "組員：＿＿＿、＿＿＿、＿＿＿、＿＿＿、＿＿＿、＿＿＿、＿＿＿", { x: 0.6, y: 4.5, w: 8.8, h: 0.35, fontSize: 14, color: "D9C3CB" });
  s.addNotes("開場：介紹組員與題目。口腔癌是台灣男性第 4 大癌症，也是最容易預防的癌症之一。");

  // 2 Agenda
  s = contentSlide("報告大綱", "開場");
  const agenda = ["器官基本功能與疾病定義", "台灣現況統計", "成因與高危險群", "症狀", "預防方法", "主要治療方式", "實例：新聞報導", "資料來源", "小組分工"];
  agenda.forEach((a, i) => {
    const col = i < 5 ? 0 : 1, row = i < 5 ? i : i - 5;
    const x = 0.5 + col * 4.6, y = 1.3 + row * 0.7;
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.5, h: 0.5, fill: { color: i % 2 ? HEX.teal : HEX.berry }, line: { color: i % 2 ? HEX.teal : HEX.berry } });
    txt(s, String(i + 1), { x, y, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: HEX.white, align: "center", valign: "middle" });
    txt(s, a, { x: x + 0.7, y, w: 3.7, h: 0.5, fontSize: 18, color: HEX.ink, valign: "middle" });
  });

  // 2b One-minute summary
  s = contentSlide("一分鐘看懂口腔癌", "開場");
  const quick = [
    [fa.FaQuestion, "是什麼", "長在嘴巴裡的癌症"], [fa.FaUsers, "多常見", "每年約 8,500 人得病"],
    [fa.FaExclamationTriangle, "誰最危險", "吃檳榔、抽菸、喝酒的人"], [fa.FaEye, "怎麼發現", "嘴破兩週沒好就要看醫生"],
    [fa.FaShieldAlt, "怎麼預防", "戒菸檳＋每 2 年免費檢查"], [fa.FaHeartbeat, "好消息", "早發現，8 成以上能治好"],
  ];
  for (let i = 0; i < quick.length; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.25 + row * 1.8;
    card(s, x, y, 2.85, 1.6, row === 0 ? HEX.rose : HEX.mint, "quick-" + i);
    await iconCircle(s, quick[i][0], x + 0.2, y + 0.2, 0.55, row === 0 ? HEX.berry : HEX.teal, HEX.white);
    txt(s, quick[i][1], { x: x + 0.9, y: y + 0.2, w: 1.8, h: 0.55, fontSize: 18, bold: true, color: row === 0 ? HEX.berry : HEX.teal, valign: "middle" });
    txt(s, quick[i][2], { x: x + 0.2, y: y + 0.9, w: 2.5, h: 0.55, fontSize: 15, color: HEX.ink, valign: "top" });
  }

  // 3 Anatomy & function
  pres.addSection({ title: "器官與定義" });
  s = contentSlide("嘴巴每天幫我們做很多事", "器官與定義");
  card(s, 0.5, 1.25, 3.6, 3.55, HEX.rose, "parts-card");
  txt(s, "嘴巴裡哪些地方會長", { x: 0.75, y: 1.4, w: 3.2, h: 0.4, fontSize: 16, bold: true, color: HEX.berry });
  txt(s, ["嘴唇", "臉頰內側（台灣最常見）", "舌頭", "舌頭下面", "牙齦（包住牙齒的肉）", "上顎（嘴巴的天花板）", "最裡面大牙的後方"].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
    { x: 0.75, y: 1.85, w: 3.2, h: 2.8, fontSize: 14, color: HEX.ink, paraSpaceAfter: 4, valign: "top" });
  const funcs = [
    [fa.FaUtensils, "咬碎食物", "牙齒和舌頭把食物磨碎"], [fa.FaArrowDown, "吞東西", "舌頭把食物推進喉嚨"],
    [fa.FaFlask, "開始消化", "飯嚼久了會變甜"], [fa.FaCommentDots, "說話", "舌頭、嘴唇一起發音"],
    [fa.FaLemon, "嚐味道", "酸甜苦鹹都靠舌頭"], [fa.FaShieldAlt, "擋細菌", "口水像一道防護牆"],
  ];
  for (let i = 0; i < funcs.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 4.4 + col * 2.6, y = 1.25 + row * 1.2;
    await iconCircle(s, funcs[i][0], x, y + 0.1, 0.6, HEX.berry, HEX.white);
    txt(s, funcs[i][1], { x: x + 0.75, y: y + 0.05, w: 1.8, h: 0.35, fontSize: 16, bold: true, color: HEX.ink });
    txt(s, funcs[i][2], { x: x + 0.75, y: y + 0.42, w: 1.8, h: 0.5, fontSize: 12, color: HEX.grey, valign: "top" });
  }
  s.addNotes("口腔不只用來吃東西，也負責說話與防禦。口腔癌手術常影響咀嚼、吞嚥與說話，因此對生活品質衝擊很大。");

  // 4 Definition
  s = contentSlide("什麼是口腔癌", "器官與定義");
  card(s, 0.5, 1.25, 9.0, 1.05, HEX.rose, "def-card");
  txt(s, [
    { text: "嘴巴內側的表皮長年被檳榔、菸、酒刺激，就像同一個地方一直被磨破，久了細胞「壞掉」開始亂長，", options: {} },
    { text: "這就是口腔癌。", options: { bold: true, color: HEX.berry } },
  ], { x: 0.75, y: 1.35, w: 8.5, h: 0.85, fontSize: 15, color: HEX.ink, valign: "middle" });
  // progression flow
  const steps = [["正常的嘴巴", HEX.mint, HEX.teal], ["警告病變", "FBE7D6", HEX.gold], ["早期癌 (第1–2期)", HEX.rose, HEX.berry], ["晚期癌 (第3–4期)", "E9C9D3", HEX.ink]];
  steps.forEach((st, i) => {
    const x = 0.5 + i * 2.32;
    card(s, x, 2.6, 2.0, 0.8, st[1], "step-" + i);
    txt(s, st[0], { x, y: 2.6, w: 2.0, h: 0.8, fontSize: 15, bold: true, color: st[2], align: "center", valign: "middle" });
    if (i < 3) s.addShape(pres.shapes.CHEVRON, { x: x + 2.06, y: 2.85, w: 0.2, h: 0.3, fill: { color: "B98A9A" }, line: { color: "B98A9A" } });
  });
  txt(s, "警告病變（癌前病變）：還不是癌，但可能變成癌，例如白斑、紅斑、嘴巴變硬；早發現就能阻止", { x: 0.5, y: 3.6, w: 9.0, h: 0.4, fontSize: 14, color: HEX.ink });
  txt(s, "分期：數字越大越嚴重；第 3、4 期代表腫瘤很大，或已跑到脖子或其他器官", { x: 0.5, y: 4.05, w: 9.0, h: 0.4, fontSize: 14, color: HEX.ink });
  txt(s, "小提醒：政府統計的「口腔癌」也包含喉嚨上段的癌症", { x: 0.5, y: 4.5, w: 9.0, h: 0.35, fontSize: 12, italic: true, color: HEX.grey });
  SRC(s, "臺北榮總 口腔癌治療準則");

  // 5 Stats
  pres.addSection({ title: "台灣現況" });
  s = contentSlide("台灣每年 8 千多人罹患口腔癌", "台灣現況");
  const stats = [["8,522", "人", "2023 年新得到\n口腔癌的人數", HEX.berry], ["3,697", "人", "2025 年因口腔癌\n過世的人數", HEX.ink], ["第 4", "名", "男性最常得的\n癌症排名", HEX.berry], ["60", "歲", "一半的病人\n不到 60 歲", HEX.teal]];
  stats.forEach((st, i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.3, 2.1, 2.6, i === 1 ? "EFE7EA" : (i === 3 ? HEX.mint : HEX.rose), "stat-" + i);
    txt(s, [{ text: st[0], options: { fontSize: 40, bold: true, color: st[3] } }, { text: " " + st[1], options: { fontSize: 16, color: st[3] } }], { x: x + 0.15, y: 1.55, w: 1.8, h: 0.9, align: "center", valign: "middle" });
    txt(s, st[2], { x: x + 0.15, y: 2.55, w: 1.8, h: 1.1, fontSize: 13, color: HEX.ink, align: "center", valign: "top" });
  });
  txt(s, "平均每天約 23 人得病、約 10 人過世；扣除人口變老的影響後，2025 年比 2024 年改善", { x: 0.5, y: 4.15, w: 9.0, h: 0.4, fontSize: 14, color: HEX.teal, bold: true });
  SRC(s, "國健署 112 年癌症登記報告（2025 公布）；衛福部 114 年死因統計（2026 公布）");
  s.addNotes("癌症登記最新為 112 年資料，死因統計最新為 114 年資料。口腔癌新發 8,522 人，比 111 年多 172 人。");

  // 6 Gender chart
  s = contentSlide("患者約 9 成是男性", "台灣現況");
  s.addChart(pres.charts.DOUGHNUT, [{ name: "112 年口腔癌新發生", labels: ["男性", "女性"], values: [7641, 881] }], {
    x: 0.5, y: 1.2, w: 4.4, h: 3.6, holeSize: 60, chartColors: [HEX.berry, "D8B4C1"],
    showPercent: true, showValue: false, showLegend: true, legendPos: "b", legendFontSize: 12, legendFontFace: "+mn-lt",
    dataLabelColor: HEX.white, dataLabelFontSize: 14, dataLabelFontFace: "+mn-lt", showTitle: false,
  });
  txt(s, "7,641", { x: 5.3, y: 1.4, w: 4.2, h: 0.8, fontSize: 44, bold: true, color: HEX.berry });
  txt(s, "2023 年新得到口腔癌的男性，每 10 個病人約 9 個是男生", { x: 5.3, y: 2.2, w: 4.2, h: 0.6, fontSize: 14, color: HEX.ink, valign: "top" });
  txt(s, [
    { text: "男性患者逾 8 成吸菸、約 7 成嚼檳榔", options: { bullet: true, breakLine: true } },
    { text: "多是中年男性，常是家裡主要賺錢的人", options: { bullet: true } },
  ], { x: 5.3, y: 3.0, w: 4.2, h: 1.2, fontSize: 14, color: HEX.ink, paraSpaceAfter: 6, valign: "top" });
  SRC(s, "國健署 112 年癌症登記報告（女性人數 = 8,522 − 7,641）；國健署新聞稿");

  // 7 Causes chart
  pres.addSection({ title: "成因與症狀" });
  s = contentSlide("檳榔、菸、酒：風險層層加乘", "成因與症狀");
  s.addChart(pres.charts.BAR, [{ name: "相對風險（倍）", labels: ["嚼檳榔", "檳榔＋吸菸", "檳榔＋菸＋酒"], values: [28, 89, 123] }], {
    x: 0.5, y: 1.2, w: 5.0, h: 3.6, barDir: "col", chartColors: ["D8B4C1", "B98A9A", HEX.berry], varyColors: true,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.ink, dataLabelFormatCode: "0\"倍\"",
    catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "得口腔癌的機會是一般人的幾倍", titleFontSize: 13, titleColor: HEX.grey, titleFontFace: "+mn-lt",
  });
  card(s, 5.8, 1.25, 3.7, 3.55, HEX.rose, "iarc-card");
  txt(s, "檳榔本身就會致癌", { x: 6.0, y: 1.4, w: 3.3, h: 0.45, fontSize: 16, bold: true, color: HEX.berry });
  txt(s, [
    { text: "世界衛生組織的癌症研究機構把檳榔列為「確定會致癌」，不加紅灰、白灰也一樣", options: { bullet: true, breakLine: true } },
    { text: "檳榔的粗纖維一直磨破嘴巴，造成反覆的傷口", options: { bullet: true, breakLine: true } },
    { text: "其他：某些病毒感染、假牙不合、不刷牙", options: { bullet: true } },
  ], { x: 6.0, y: 1.9, w: 3.3, h: 2.8, fontSize: 13, color: HEX.ink, paraSpaceAfter: 6, valign: "top" });
  SRC(s, "葛應欽等（1995）台灣口腔癌與嚼檳榔、吸菸、酒精之關聯，引自國軍衛教資料；國健署");
  s.addNotes("28/89/123 倍出自 1995 年台灣研究，屬經典但較舊的數據；國健署 2022 年另指出嚼檳榔加菸酒，上呼吸消化道癌風險高 10.5 倍。");

  // 8 High risk
  s = contentSlide("誰是高危險群", "成因與症狀");
  const risks = [
    [fa.FaSmoking, "30 歲以上吸菸或嚼檳榔者", "含已戒檳榔者"], [fa.FaUsers, "18 歲以上嚼檳榔原住民", "含已戒檳榔者"],
    [fa.FaMale, "中年男性", "每 10 個病人約 9 個是男生"], [fa.FaWineGlassAlt, "菸、酒、檳榔多重使用", "風險加乘最高"],
    [fa.FaTruck, "需長時間提神的勞動工作者", "如貨運司機、建築工"], [fa.FaNotesMedical, "嘴巴已經有異常的人", "白斑、紅斑、嘴巴內側變硬"],
  ];
  for (let i = 0; i < risks.length; i++) {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6, y = 1.25 + row * 1.2;
    card(s, x, y, 4.4, 1.0, i < 2 ? HEX.rose : "F4F1F2", "risk-" + i);
    await iconCircle(s, risks[i][0], x + 0.2, y + 0.18, 0.64, i < 2 ? HEX.berry : HEX.grey, HEX.white);
    txt(s, risks[i][1], { x: x + 1.0, y: y + 0.12, w: 3.3, h: 0.42, fontSize: 15, bold: true, color: HEX.ink });
    txt(s, risks[i][2], { x: x + 1.0, y: y + 0.54, w: 3.3, h: 0.35, fontSize: 12, color: HEX.grey });
  }
  SRC(s, "國健署口腔黏膜檢查補助資格；國健署 112 年癌症登記報告");

  // 9 Symptoms
  s = contentSlide("五大警訊：破、斑、硬、突、腫", "成因與症狀");
  const sx = [["破", "嘴破超過兩個禮拜還沒好"], ["斑", "出現擦不掉的白色或紅色斑塊"], ["硬", "嘴巴內側變硬，嘴巴張不開"], ["突", "長出凹凸不平的小肉塊"], ["腫", "臉或脖子腫起來、左右不一樣"]];
  sx.forEach((t, i) => {
    const x = 0.5 + i * 1.84;
    card(s, x, 1.25, 1.66, 2.55, i === 0 ? HEX.berry : HEX.rose, "sx-" + i);
    txt(s, t[0], { x, y: 1.4, w: 1.66, h: 0.95, fontSize: 48, bold: true, color: i === 0 ? HEX.white : HEX.berry, align: "center", valign: "middle" });
    txt(s, t[1], { x: x + 0.15, y: 2.45, w: 1.36, h: 1.2, fontSize: 13, color: i === 0 ? HEX.white : HEX.ink, align: "center", valign: "top" });
  });
  txt(s, "其他：嘴巴流血或麻麻的、牙齒突然鬆動、吞東西或說話困難、一直變瘦", { x: 0.5, y: 3.95, w: 9.0, h: 0.4, fontSize: 14, color: HEX.ink });
  txt(s, "早期常常不會痛，很多人以為只是嘴破，自己擦藥而拖延", { x: 0.5, y: 4.38, w: 9.0, h: 0.4, fontSize: 14, bold: true, color: HEX.berry });
  SRC(s, "衛福部國健署「破斑硬突腫」口腔癌警訊");

  // 10 Prevention
  pres.addSection({ title: "預防與治療" });
  s = contentSlide("預防：拒菸檳、定期篩、早治療", "預防與治療");
  const prev = [[fa.FaBan, "拒菸檳", "不吃檳榔、不抽菸、少喝酒；戒檳榔 10 年，風險可降約二到五成", HEX.berry], [fa.FaSearch, "定期篩", "政府免費幫你檢查嘴巴，每 2 年 1 次，只要 3 分鐘、不會痛", HEX.teal], [fa.FaHospital, "早治療", "檢查有異常，就到大醫院切一小塊化驗，早點治療", HEX.ink]];
  for (let i = 0; i < prev.length; i++) {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.25, 2.85, 2.0, i === 1 ? HEX.mint : HEX.rose, "prev-" + i);
    await iconCircle(s, prev[i][0], x + 0.2, 1.42, 0.6, prev[i][3], HEX.white);
    txt(s, prev[i][1], { x: x + 0.95, y: 1.42, w: 1.8, h: 0.6, fontSize: 20, bold: true, color: prev[i][3], valign: "middle" });
    txt(s, prev[i][2], { x: x + 0.2, y: 2.15, w: 2.5, h: 1.0, fontSize: 13, color: HEX.ink, valign: "top" });
  }
  txt(s, [{ text: "-26%", options: { fontSize: 36, bold: true, color: HEX.teal } }, { text: "  因口腔癌過世的風險", options: { fontSize: 15, color: HEX.ink } }], { x: 0.5, y: 3.45, w: 4.4, h: 0.7, valign: "middle" });
  txt(s, [{ text: "-21%", options: { fontSize: 36, bold: true, color: HEX.teal } }, { text: "  拖到很嚴重才發現的風險", options: { fontSize: 15, color: HEX.ink } }], { x: 5.1, y: 3.45, w: 4.4, h: 0.7, valign: "middle" });
  txt(s, "定期檢查嘴巴的效果；檢查找到的問題，約 8 成是還沒變成癌或很早期的", { x: 0.5, y: 4.2, w: 9.0, h: 0.5, fontSize: 13, color: HEX.grey });
  SRC(s, "衛福部「定期口腔黏膜檢查 降低 26% 死亡風險」（引臺大研究）");

  // 11 Screening & policy
  s = contentSlide("免費嘴巴檢查與政府政策", "預防與治療");
  s.addTable([
    [{ text: "項目", options: { bold: true, color: HEX.white, fill: { color: HEX.berry } } }, { text: "內容", options: { bold: true, color: HEX.white, fill: { color: HEX.berry } } }],
    ["誰可以做", "30 歲以上有吃檳榔（含已戒）或抽菸的人；18 歲以上有吃檳榔（含已戒）的原住民"],
    ["多久一次", "每 2 年 1 次，免費"],
    ["去哪裡做", "帶健保卡到有健保的牙科或耳鼻喉科"],
    ["有沒有用", "每 2 年約找出近 3,000 個警告病變、1,000 多個口腔癌病人"],
  ], { x: 0.5, y: 1.25, w: 5.2, colW: [1.2, 4.0], fontSize: 12, fontFace: FONT, color: HEX.ink, border: { type: "solid", pt: 0.5, color: "E2CDD5" }, fill: { color: HEX.white }, valign: "middle", rowH: 0.55 });
  card(s, 6.0, 1.25, 3.5, 3.5, HEX.mint, "policy-card");
  txt(s, "政策動態", { x: 6.2, y: 1.38, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: HEX.teal });
  txt(s, [
    { text: "2025 年起放寬免費癌症檢查；2026 年起可免費檢查 6 種癌症", options: { bullet: true, breakLine: true } },
    { text: "2024 年免費癌症檢查約 481 萬人次，找出約 5 萬個警告病變", options: { bullet: true, breakLine: true } },
    { text: "管理檳榔的法律草案 2024 年底提出，2026 年仍未送立法院，監察院已批評", options: { bullet: true } },
  ], { x: 6.2, y: 1.85, w: 3.1, h: 2.8, fontSize: 12, color: HEX.ink, paraSpaceAfter: 6, valign: "top" });
  SRC(s, "國健署新聞稿；IQ 健康網（2026 癌篩）；聯合報、工商時報（2026）");

  // 12 Treatment
  s = contentSlide("主要治療方式", "預防與治療");
  const tx = [
    [fa.FaCut, "開刀", "最主要的治療；切掉腫瘤，再從手或腳取皮肉補回去"],
    [fa.FaRadiation, "電療", "用高能量射線照癌細胞；會口乾、嘴破"],
    [fa.FaSyringe, "電療＋化療", "容易復發或太大切不掉時，兩種一起做"],
    [fa.FaPills, "化療", "用藥物殺全身的癌細胞；會想吐、掉頭髮"],
    [fa.FaCrosshairs, "標靶治療", "像導彈一樣，專打癌細胞的特定目標"],
    [fa.FaShieldVirus, "免疫治療", "拆穿癌細胞的偽裝，讓身體自己去打"],
  ];
  for (let i = 0; i < tx.length; i++) {
    const col = i % 3, row = Math.floor(i / 3);
    const x = 0.5 + col * 3.05, y = 1.25 + row * 1.75;
    card(s, x, y, 2.85, 1.55, row === 0 ? HEX.rose : "F4F1F2", "tx-" + i);
    await iconCircle(s, tx[i][0], x + 0.18, y + 0.18, 0.55, row === 0 ? HEX.berry : HEX.grey, HEX.white);
    txt(s, tx[i][1], { x: x + 0.85, y: y + 0.18, w: 1.9, h: 0.55, fontSize: 16, bold: true, color: HEX.ink, valign: "middle" });
    txt(s, tx[i][2], { x: x + 0.18, y: y + 0.82, w: 2.5, h: 0.65, fontSize: 12, color: HEX.ink, valign: "top" });
  }
  SRC(s, "臺北榮總口腔癌治療準則；臺大醫院健康電子報；聯合報（免疫治療給付）");

  // 13 Survival chart
  s = contentSlide("越早發現，越有機會活下來", "預防與治療");
  s.addChart(pres.charts.BAR, [{ name: "五年存活率（%）", labels: ["第 1 期", "第 2 期", "第 3 期", "第 4 期"], values: [85, 70, 60, 30] }], {
    x: 0.5, y: 1.2, w: 5.4, h: 3.6, barDir: "col", chartColors: [HEX.teal, "5FA39C", "B98A9A", HEX.berry], varyColors: true,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.ink, dataLabelFormatCode: "0\"人\"",
    catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "每 100 位病人中，活超過 5 年的人數（大約值）", titleFontSize: 13, titleColor: HEX.grey, titleFontFace: "+mn-lt",
  });
  txt(s, "8 成以上", { x: 6.2, y: 1.4, w: 3.3, h: 0.8, fontSize: 36, bold: true, color: HEX.teal });
  txt(s, "早期發現的病人，活超過 5 年的比例（國健署相關研究）", { x: 6.2, y: 2.2, w: 3.3, h: 0.6, fontSize: 14, color: HEX.ink, valign: "top" });
  txt(s, "但超過 7 成的頭頸部癌症病人，被發現時已經比較嚴重（醫師，2026）", { x: 6.2, y: 3.1, w: 3.3, h: 1.0, fontSize: 14, color: HEX.berry, bold: true, valign: "top" });
  SRC(s, "NOWnews、TVBS 醫師訪談（各院數字略有差異，第 1 期取 80%–90% 中間值）");

  // 14 News
  pres.addSection({ title: "實例" });
  s = contentSlide("新聞實例", "實例");
  const news = [
    ["2026/06", "監察院批評行政院", "管理檳榔的法律拖了很久，到現在還沒送立法院", "聯合報"],
    ["2026/03", "45 歲男長期菸檳", "醫師：頭頸部癌症以口腔癌最多，逾 7 成發現時已較嚴重", "TVBS 健康 2+1"],
    ["2025/09", "藝人江國賓口腔腫瘤", "年輕時抽菸吃檳榔，化驗是良性（不是癌），醫師提醒要一輩子定期檢查", "TVBS 健康 2+1"],
    ["2020", "嘴破 2 個月沒好", "36 歲男子長期吃檳榔提神，嘴破自己擦藥，看醫生時已是最嚴重的第 4 期", "TVBS 新聞"],
  ];
  news.forEach((n, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6, y = 1.25 + row * 1.85;
    card(s, x, y, 4.4, 1.7, HEX.rose, "news-" + i);
    txt(s, n[0], { x: x + 0.2, y: y + 0.12, w: 1.2, h: 0.3, fontSize: 11, bold: true, color: HEX.teal });
    txt(s, n[3], { x: x + 2.0, y: y + 0.12, w: 2.2, h: 0.3, fontSize: 11, color: HEX.grey, align: "right" });
    txt(s, n[1], { x: x + 0.2, y: y + 0.45, w: 4.0, h: 0.4, fontSize: 16, bold: true, color: HEX.berry });
    txt(s, n[2], { x: x + 0.2, y: y + 0.88, w: 4.0, h: 0.75, fontSize: 13, color: HEX.ink, valign: "top" });
  });
  s.addNotes("江國賓案例部分媒體標題寫「口腔癌」，但實際切片為良性腫瘤，報告時請正確說明。影片：可到國健署 YouTube 頻道搜尋「口腔癌篩檢」挑一部 3 分鐘內的影片播放。");

  // 15 Division of labour
  pres.addSection({ title: "分工與來源" });
  s = contentSlide("小組分工表", "分工與來源");
  const hdr = (t) => ({ text: t, options: { bold: true, color: HEX.white, fill: { color: HEX.berry } } });
  s.addTable([
    [hdr("組員"), hdr("負責項目"), hdr("具體工作")],
    ["組員 1（組長）", "統籌、器官功能與疾病定義", "分配進度、彙整書面報告；撰寫構造、功能、定義"],
    ["組員 2", "台灣現況統計", "查核癌症登記與死因統計最新數字"],
    ["組員 3", "成因與高危險群", "整理檳榔、菸、酒等危險因子"],
    ["組員 4", "症狀", "說明「破斑硬突腫」並找示意圖"],
    ["組員 5", "預防方法", "戒菸檳、公費篩檢與政策"],
    ["組員 6", "主要治療方式", "手術、放化療、標靶、免疫與存活率"],
    ["組員 7", "實例、資料來源", "挑選新聞影片、核對連結與參考格式"],
  ], { x: 0.5, y: 1.2, w: 9.0, colW: [1.7, 2.6, 4.7], fontSize: 12, fontFace: FONT, color: HEX.ink, border: { type: "solid", pt: 0.5, color: "E2CDD5" }, fill: { color: HEX.white }, valign: "middle", rowH: 0.43 });
  txt(s, "請換成真實姓名並依老師規定格式調整；每人另交個人心得 200–300 字", { x: 0.5, y: 4.85, w: 9.0, h: 0.3, fontSize: 10, color: HEX.grey });

  // 16 Sources
  s = contentSlide("資料來源", "分工與來源");
  const srcL = ["衛生福利部（2026）。114 年國人死因統計結果", "衛生福利部（2025）。113 年國人死因統計結果", "國民健康署（2025）。112 年國人癌症登記資料分析結果", "國民健康署（2024）。111 年癌症登記報告記者會簡報", "國民健康署。口腔癌 2 年 1 篩、114 年起擴大癌症篩檢", "衛福部。定期口腔黏膜檢查降低 26% 死亡風險；破斑硬突腫"];
  const srcR = ["臺北榮民總醫院。口腔癌治療準則", "臺大醫院健康電子報 236 期。頭頸癌的治療新趨勢與挑戰", "葛應欽等（1995）。口腔癌與嚼檳榔、吸菸、酒精之關聯", "聯合報（2026）；工商時報（2026）：檳榔防制專法", "TVBS（2020、2025、2026）口腔癌相關報導", "鏡週刊（2023）；聯合報：健保免疫治療給付"];
  txt(s, "政府機關", { x: 0.5, y: 1.15, w: 4.4, h: 0.35, fontSize: 14, bold: true, color: HEX.berry });
  txt(s, srcL.map((t, i, a) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < a.length - 1 } })), { x: 0.5, y: 1.5, w: 4.4, h: 3.3, fontSize: 13, color: HEX.ink, paraSpaceAfter: 8, valign: "top" });
  txt(s, "醫學中心、研究與新聞", { x: 5.1, y: 1.15, w: 4.4, h: 0.35, fontSize: 14, bold: true, color: HEX.berry });
  txt(s, srcR.map((t, i, a) => ({ text: t, options: { bullet: { type: "number", startAt: 7 }, breakLine: i < a.length - 1 } })), { x: 5.1, y: 1.5, w: 4.4, h: 3.3, fontSize: 13, color: HEX.ink, paraSpaceAfter: 8, valign: "top" });
  txt(s, "完整網址見書面報告「資料來源」一節", { x: 0.5, y: 4.85, w: 9.0, h: 0.3, fontSize: 10, color: HEX.grey });

  // 17 Closing
  s = pres.addSlide({ masterName: "DARK", sectionTitle: "分工與來源" });
  s.addText("拒菸檳、定期篩、早治療", { placeholder: "title" });
  s.addText("口腔癌是最容易預防的癌症之一 —— 提醒身邊有菸檳習慣的家人，每 2 年做一次免費嘴巴檢查", { placeholder: "body" });
  await iconCircle(s, fa.FaHandHoldingHeart, 7.9, 0.5, 1.2, HEX.teal, HEX.white);
  txt(s, "謝謝聆聽．Q & A", { x: 0.6, y: 4.5, w: 8.8, h: 0.4, fontSize: 16, color: "D9C3CB" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
