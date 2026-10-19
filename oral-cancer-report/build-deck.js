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

  // 3 Anatomy & function
  pres.addSection({ title: "器官與定義" });
  s = contentSlide("口腔：消化與呼吸的入口", "器官與定義");
  card(s, 0.5, 1.25, 3.6, 3.55, HEX.rose, "parts-card");
  txt(s, "口腔包含哪些部位", { x: 0.75, y: 1.4, w: 3.2, h: 0.4, fontSize: 16, bold: true, color: HEX.berry });
  txt(s, ["唇", "頰黏膜（台灣最常見部位之一）", "舌頭前 2/3", "口底", "上、下牙齦", "硬顎", "臼齒後三角"].map((t, i, a) => ({ text: t, options: { bullet: true, breakLine: i < a.length - 1 } })),
    { x: 0.75, y: 1.85, w: 3.2, h: 2.8, fontSize: 14, color: HEX.ink, paraSpaceAfter: 4, valign: "top" });
  const funcs = [
    [fa.FaUtensils, "咀嚼", "磨碎食物、混合唾液"], [fa.FaArrowDown, "吞嚥", "舌頭推送食團入咽"],
    [fa.FaFlask, "消化", "唾液澱粉酶分解澱粉"], [fa.FaCommentDots, "語言", "舌、唇、齒配合發音"],
    [fa.FaLemon, "味覺", "味蕾感受五種味道"], [fa.FaShieldAlt, "防禦", "黏膜與唾液阻擋細菌"],
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
    { text: "口腔黏膜細胞受長期刺激（檳榔、菸、酒）產生基因突變，不受控制地增生並侵犯周圍組織，甚至轉移到頸部淋巴結。", options: { breakLine: true } },
    { text: "約 9 成為「鱗狀上皮細胞癌」。", options: { bold: true, color: HEX.berry } },
  ], { x: 0.75, y: 1.35, w: 8.5, h: 0.85, fontSize: 15, color: HEX.ink, valign: "middle" });
  // progression flow
  const steps = [["正常黏膜", HEX.mint, HEX.teal], ["癌前病變", "FBE7D6", HEX.gold], ["早期癌 (1–2 期)", HEX.rose, HEX.berry], ["晚期癌 (3–4 期)", "E9C9D3", HEX.ink]];
  steps.forEach((st, i) => {
    const x = 0.5 + i * 2.32;
    card(s, x, 2.6, 2.0, 0.8, st[1], "step-" + i);
    txt(s, st[0], { x, y: 2.6, w: 2.0, h: 0.8, fontSize: 15, bold: true, color: st[2], align: "center", valign: "middle" });
    if (i < 3) s.addShape(pres.shapes.CHEVRON, { x: x + 2.06, y: 2.85, w: 0.2, h: 0.3, fill: { color: "B98A9A" }, line: { color: "B98A9A" } });
  });
  txt(s, "癌前病變：白斑、紅斑、口腔黏膜下纖維化、疣狀增生、扁平苔癬 —— 及早發現就能阻止惡化", { x: 0.5, y: 3.6, w: 9.0, h: 0.4, fontSize: 14, color: HEX.ink });
  txt(s, "分期：依 TNM（腫瘤大小 T、淋巴結 N、遠處轉移 M）分為 0–4 期", { x: 0.5, y: 4.05, w: 9.0, h: 0.4, fontSize: 14, color: HEX.ink });
  txt(s, "統計範圍：國健署癌症登記的「口腔癌」含口咽、下咽", { x: 0.5, y: 4.5, w: 9.0, h: 0.35, fontSize: 12, italic: true, color: HEX.grey });
  SRC(s, "臺北榮總 口腔癌治療準則");

  // 5 Stats
  pres.addSection({ title: "台灣現況" });
  s = contentSlide("台灣每年 8 千多人罹患口腔癌", "台灣現況");
  const stats = [["8,522", "人", "112 年新發生人數\n（全癌第 6 位）", HEX.berry], ["3,697", "人", "114 年死亡人數\n（癌症死因第 6 位）", HEX.ink], ["第 4", "位", "男性十大癌症\n（男性 7,641 人）", HEX.berry], ["60", "歲", "發生年齡中位數\n（112 年）", HEX.teal]];
  stats.forEach((st, i) => {
    const x = 0.5 + i * 2.3;
    card(s, x, 1.3, 2.1, 2.6, i === 1 ? "EFE7EA" : (i === 3 ? HEX.mint : HEX.rose), "stat-" + i);
    txt(s, [{ text: st[0], options: { fontSize: 40, bold: true, color: st[3] } }, { text: " " + st[1], options: { fontSize: 16, color: st[3] } }], { x: x + 0.15, y: 1.55, w: 1.8, h: 0.9, align: "center", valign: "middle" });
    txt(s, st[2], { x: x + 0.15, y: 2.55, w: 1.8, h: 1.1, fontSize: 13, color: HEX.ink, align: "center", valign: "top" });
  });
  txt(s, "114 年口腔癌標準化死亡率較 113 年下降；每十萬人口死亡率 15.8 人", { x: 0.5, y: 4.15, w: 9.0, h: 0.4, fontSize: 14, color: HEX.teal, bold: true });
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
  txt(s, "112 年男性口腔癌新發生人數，較 111 年增加 169 人", { x: 5.3, y: 2.2, w: 4.2, h: 0.6, fontSize: 14, color: HEX.ink, valign: "top" });
  txt(s, [
    { text: "男性患者逾 8 成吸菸、約 7 成嚼檳榔", options: { bullet: true, breakLine: true } },
    { text: "好發於中壯年，常是家庭經濟支柱", options: { bullet: true } },
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
    showTitle: true, title: "罹患口腔癌風險（與不使用者相比）", titleFontSize: 13, titleColor: HEX.grey, titleFontFace: "+mn-lt",
  });
  card(s, 5.8, 1.25, 3.7, 3.55, HEX.rose, "iarc-card");
  txt(s, "檳榔本身就是一級致癌物", { x: 6.0, y: 1.4, w: 3.3, h: 0.45, fontSize: 16, bold: true, color: HEX.berry });
  txt(s, [
    { text: "IARC 將檳榔子列為第一級致癌物，不加紅灰、白灰、荖葉一樣致癌", options: { bullet: true, breakLine: true } },
    { text: "檳榔纖維反覆摩擦黏膜造成傷口", options: { bullet: true, breakLine: true } },
    { text: "其他：HPV（口咽癌）、不合適假牙、口腔衛生不良", options: { bullet: true } },
  ], { x: 6.0, y: 1.9, w: 3.3, h: 2.8, fontSize: 13, color: HEX.ink, paraSpaceAfter: 6, valign: "top" });
  SRC(s, "葛應欽等（1995）台灣口腔癌與嚼檳榔、吸菸、酒精之關聯，引自國軍衛教資料；國健署");
  s.addNotes("28/89/123 倍出自 1995 年台灣研究，屬經典但較舊的數據；國健署 2022 年另指出嚼檳榔加菸酒，上呼吸消化道癌風險高 10.5 倍。");

  // 8 High risk
  s = contentSlide("誰是高危險群", "成因與症狀");
  const risks = [
    [fa.FaSmoking, "30 歲以上吸菸或嚼檳榔者", "含已戒檳榔者"], [fa.FaUsers, "18 歲以上嚼檳榔原住民", "含已戒檳榔者"],
    [fa.FaMale, "中壯年男性", "約 9 成患者為男性"], [fa.FaWineGlassAlt, "菸、酒、檳榔多重使用", "風險加乘最高"],
    [fa.FaTruck, "需長時間提神的勞動工作者", "如貨運司機、建築工"], [fa.FaNotesMedical, "已有癌前病變或頭頸癌病史", "白斑、紅斑、黏膜下纖維化"],
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
  const sx = [["破", "黏膜破皮、潰瘍超過兩週未癒合"], ["斑", "出現不明原因的白斑或紅斑"], ["硬", "黏膜變硬變緊，嘴巴張不開"], ["突", "表面不規則突起（乳突或疣狀）"], ["腫", "臉或頸部腫塊、兩側不對稱"]];
  sx.forEach((t, i) => {
    const x = 0.5 + i * 1.84;
    card(s, x, 1.25, 1.66, 2.55, i === 0 ? HEX.berry : HEX.rose, "sx-" + i);
    txt(s, t[0], { x, y: 1.4, w: 1.66, h: 0.95, fontSize: 48, bold: true, color: i === 0 ? HEX.white : HEX.berry, align: "center", valign: "middle" });
    txt(s, t[1], { x: x + 0.15, y: 2.45, w: 1.36, h: 1.2, fontSize: 13, color: i === 0 ? HEX.white : HEX.ink, align: "center", valign: "top" });
  });
  txt(s, "其他症狀：口腔出血或麻木、牙齒鬆動、吞嚥或說話困難、持續口臭、體重減輕", { x: 0.5, y: 3.95, w: 9.0, h: 0.4, fontSize: 14, color: HEX.ink });
  txt(s, "早期常不痛，容易被當成一般嘴破而延誤就醫", { x: 0.5, y: 4.38, w: 9.0, h: 0.4, fontSize: 14, bold: true, color: HEX.berry });
  SRC(s, "衛福部國健署「破斑硬突腫」口腔癌警訊");

  // 10 Prevention
  pres.addSection({ title: "預防與治療" });
  s = contentSlide("預防：拒菸檳、定期篩、早治療", "預防與治療");
  const prev = [[fa.FaBan, "拒菸檳", "不嚼檳榔、不吸菸、少喝酒；戒檳 10 年風險可降 17%–51%", HEX.berry], [fa.FaSearch, "定期篩", "公費口腔黏膜檢查，每 2 年 1 次，約 3 分鐘", HEX.teal], [fa.FaHospital, "早治療", "篩檢陽性就轉介切片確診，及早治療", HEX.ink]];
  for (let i = 0; i < prev.length; i++) {
    const x = 0.5 + i * 3.05;
    card(s, x, 1.25, 2.85, 2.0, i === 1 ? HEX.mint : HEX.rose, "prev-" + i);
    await iconCircle(s, prev[i][0], x + 0.2, 1.42, 0.6, prev[i][3], HEX.white);
    txt(s, prev[i][1], { x: x + 0.95, y: 1.42, w: 1.8, h: 0.6, fontSize: 20, bold: true, color: prev[i][3], valign: "middle" });
    txt(s, prev[i][2], { x: x + 0.2, y: 2.15, w: 2.5, h: 1.0, fontSize: 13, color: HEX.ink, valign: "top" });
  }
  txt(s, [{ text: "-26%", options: { fontSize: 36, bold: true, color: HEX.teal } }, { text: "  口腔癌死亡風險", options: { fontSize: 15, color: HEX.ink } }], { x: 0.5, y: 3.45, w: 4.4, h: 0.7, valign: "middle" });
  txt(s, [{ text: "-21%", options: { fontSize: 36, bold: true, color: HEX.teal } }, { text: "  晚期口腔癌發生風險", options: { fontSize: 15, color: HEX.ink } }], { x: 5.1, y: 3.45, w: 4.4, h: 0.7, valign: "middle" });
  txt(s, "定期口腔黏膜檢查的成效；篩檢發現的病變約 77% 為癌前病變或 0–1 期早期癌", { x: 0.5, y: 4.2, w: 9.0, h: 0.5, fontSize: 13, color: HEX.grey });
  SRC(s, "衛福部「定期口腔黏膜檢查 降低 26% 死亡風險」（引臺大研究）");

  // 11 Screening & policy
  s = contentSlide("公費口腔黏膜檢查與最新政策", "預防與治療");
  s.addTable([
    [{ text: "項目", options: { bold: true, color: HEX.white, fill: { color: HEX.berry } } }, { text: "內容", options: { bold: true, color: HEX.white, fill: { color: HEX.berry } } }],
    ["補助對象", "30 歲以上嚼檳榔（含已戒）或吸菸者；18 歲以上嚼檳榔（含已戒）原住民"],
    ["頻率", "每 2 年 1 次，免費"],
    ["地點", "帶健保卡到健保特約牙科或耳鼻喉科"],
    ["成果", "每輪約找出近 3,000 名癌前病變、1,000 多名口腔癌"],
  ], { x: 0.5, y: 1.25, w: 5.2, colW: [1.2, 4.0], fontSize: 12, fontFace: FONT, color: HEX.ink, border: { type: "solid", pt: 0.5, color: "E2CDD5" }, fill: { color: HEX.white }, valign: "middle", rowH: 0.55 });
  card(s, 6.0, 1.25, 3.5, 3.5, HEX.mint, "policy-card");
  txt(s, "政策動態", { x: 6.2, y: 1.38, w: 3.1, h: 0.4, fontSize: 16, bold: true, color: HEX.teal });
  txt(s, [
    { text: "114 年起擴大癌篩；115 年元旦納入胃癌成「六癌篩檢」", options: { bullet: true, breakLine: true } },
    { text: "113 年五癌篩檢約 481.2 萬人次，找出 50,712 例癌前病變", options: { bullet: true, breakLine: true } },
    { text: "《檳榔健康危害防制法》草案 113/12 預告，2026 年仍未送立法院，監察院已糾正", options: { bullet: true } },
  ], { x: 6.2, y: 1.85, w: 3.1, h: 2.8, fontSize: 12, color: HEX.ink, paraSpaceAfter: 6, valign: "top" });
  SRC(s, "國健署新聞稿；IQ 健康網（2026 癌篩）；聯合報、工商時報（2026）");

  // 12 Treatment
  s = contentSlide("主要治療方式", "預防與治療");
  const tx = [
    [fa.FaCut, "手術切除", "首選；切除腫瘤＋頸部淋巴結廓清，以皮瓣重建"],
    [fa.FaRadiation, "放射治療", "術後輔助或無法手術者；副作用：口乾、黏膜炎"],
    [fa.FaSyringe, "同步化放療", "術後高復發風險或局部晚期；常用順鉑"],
    [fa.FaPills, "化學治療", "晚期、復發或轉移的全身治療"],
    [fa.FaCrosshairs, "標靶治療", "抗 EGFR 藥物 cetuximab，合併化放療"],
    [fa.FaShieldVirus, "免疫治療", "PD-1 抑制劑；健保第一線以 CPS ≥ 20 為主"],
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
  s = contentSlide("越早發現，存活率越高", "預防與治療");
  s.addChart(pres.charts.BAR, [{ name: "五年存活率（%）", labels: ["第 1 期", "第 2 期", "第 3 期", "第 4 期"], values: [85, 70, 60, 30] }], {
    x: 0.5, y: 1.2, w: 5.4, h: 3.6, barDir: "col", chartColors: [HEX.teal, "5FA39C", "B98A9A", HEX.berry], varyColors: true,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 14, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.ink, dataLabelFormatCode: "0\"%\"",
    catAxisLabelColor: HEX.ink, catAxisLabelFontSize: 13, catAxisLabelFontFace: "+mn-lt", valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100,
    valGridLine: { style: "none" }, catGridLine: { style: "none" }, showLegend: false,
    showTitle: true, title: "口腔癌分期五年存活率（醫師引述概估值）", titleFontSize: 13, titleColor: HEX.grey, titleFontFace: "+mn-lt",
  });
  txt(s, "8 成以上", { x: 6.2, y: 1.4, w: 3.3, h: 0.8, fontSize: 36, bold: true, color: HEX.teal });
  txt(s, "早期口腔癌五年存活率（國健署相關研究）", { x: 6.2, y: 2.2, w: 3.3, h: 0.6, fontSize: 14, color: HEX.ink, valign: "top" });
  txt(s, "頭頸癌超過 7 成確診時已是第三期以上（新竹臺大分院醫師，2026）", { x: 6.2, y: 3.1, w: 3.3, h: 1.0, fontSize: 14, color: HEX.berry, bold: true, valign: "top" });
  SRC(s, "NOWnews、TVBS 醫師訪談（各院數字略有差異，第 1 期取 80%–90% 中間值）");

  // 14 News
  pres.addSection({ title: "實例" });
  s = contentSlide("新聞實例", "實例");
  const news = [
    ["2026/06", "監察院糾正行政院", "《檳榔健康危害防制法》草案延宕近 5 年仍未送立法院", "聯合報"],
    ["2026/03", "45 歲男長期菸檳", "醫師：頭頸癌以口腔癌最多，逾 7 成確診已第三期以上", "TVBS 健康 2+1"],
    ["2025/09", "藝人江國賓口腔腫瘤", "年輕時吸菸嚼檳榔，切片為良性並切除；醫師提醒需終生追蹤", "TVBS 健康 2+1"],
    ["2020", "嘴破 2 個月沒好", "新竹 36 歲男子長期嚼檳提神，自行擦藥，就醫已第四期", "TVBS 新聞"],
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
  s.addText("口腔癌是最容易預防的癌症之一 —— 提醒身邊有菸檳習慣的家人，每 2 年做一次口腔黏膜檢查", { placeholder: "body" });
  await iconCircle(s, fa.FaHandHoldingHeart, 7.9, 0.5, 1.2, HEX.teal, HEX.white);
  txt(s, "謝謝聆聽．Q & A", { x: 0.6, y: 4.5, w: 8.8, h: 0.4, fontSize: 16, color: "D9C3CB" });

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
