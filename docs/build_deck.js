// Vigia — Telemetria de tanques: technical deck (PT) for Sonangol Distribuidora
const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const md = require("react-icons/md");
const { applyTheme } = require("/Users/airesmiguens/.claude/skills/synced/29b3818e-d9fc-469f-9646-34b8882402b7_46c5f86c-cd11-44a6-aa6e-a99d2ebfb779/pptx/scripts/apply_theme.js");

const OUT = process.argv[2] || "Vigia_Telemetria_Tanques.pptx";

const HEX = {
  navy: "14213D", grey: "4A5568", white: "FFFFFF", panel: "F4F5F7",
  violet: "5B4BC4", amber: "E9A23B", red: "B42318", green: "1E7A46",
  water: "4F8FD6", amberDark: "8A5A00", hair: "D5DAE1", lilac: "A99CF0", soil: "EFE8DA",
};
const THEME = {
  name: "Vigia",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: HEX.navy, lt1: HEX.white, dk2: HEX.grey, lt2: HEX.panel,
    accent1: HEX.violet, accent2: HEX.amber, accent3: HEX.red, accent4: HEX.green,
    accent5: HEX.water, accent6: HEX.amberDark, hlink: HEX.violet, folHlink: HEX.grey,
  },
};
const MONO = "Courier New";
const LOGO_W = "image/png;base64," + require("fs").readFileSync(require("path").join(__dirname, "../assets/sonangol_white.png")).toString("base64");
const LOGO_RATIO = 3040 / 854;

async function iconData(name, hex) {
  const Comp = md[name];
  if (!Comp) throw new Error("missing icon " + name);
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in
  pres.title = "Vigia — Telemetria remota de tanques";
  pres.author = "Aires Augusto Miguêns";
  pres.subject = "Arquitetura técnica de telemetria de tanques para a Sonangol Distribuidora";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const S = pres.ShapeType;

  // ---------- layouts ----------
  pres.defineSlideMaster({
    title: "VIGIA_DARK",
    background: { color: C.text1 },
    objects: [
      { placeholder: { options: { name: "eyebrow", type: "body", x: 0.6, y: 0.5, w: 12.1, h: 0.4, fontSize: 13, bold: true, color: C.accent2, fontFace: MONO, charSpacing: 3, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.95, w: 12.1, h: 1.0, fontSize: 40, bold: true, color: C.background1, margin: 0, valign: "top", align: "left" }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "VIGIA_CONTENT",
    background: { color: C.background1 },
    margin: [0.5, 0.6, 0.6, 0.6],
    objects: [
      { placeholder: { options: { name: "eyebrow", type: "body", x: 0.6, y: 0.38, w: 12.1, h: 0.32, fontSize: 12, bold: true, color: C.accent1, fontFace: MONO, charSpacing: 3, margin: 0 }, text: "" } },
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.7, w: 12.1, h: 0.75, fontSize: 34, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
      { text: { text: "Vigia · Telemetria remota de tanques · preparado para a Sonangol Distribuidora", options: { x: 0.6, y: 7.0, w: 10, h: 0.3, fontSize: 10, color: C.text2, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.55, h: 0.3, fontSize: 10, color: C.text2, align: "right", margin: 0 },
  });

  // ---------- icons ----------
  const ICONS = {};
  const need = [
    ["MdSensors", "white"], ["MdRouter", "white"], ["MdCellTower", "white"], ["MdCloud", "white"], ["MdPhoneIphone", "white"],
    ["MdLock", "white"], ["MdKey", "white"], ["MdPeople", "white"], ["MdBolt", "white"], ["MdSignalCellularAlt", "white"],
    ["MdWarning", "white"], ["MdStorage", "white"], ["MdLocalGasStation", "navy"], ["MdSchedule", "white"], ["MdLocalShipping", "white"],
    ["MdSearch", "white"], ["MdLink", "navy"], ["MdBatteryChargingFull", "white"], ["MdSettingsInputComponent", "white"],
    ["MdPrecisionManufacturing", "white"], ["MdVisibility", "navy"], ["MdWifiOff", "white"], ["MdPublic", "white"],
    ["MdTrendingDown", "white"], ["MdTrendingUp", "white"], ["MdReportProblem", "white"], ["MdHub", "white"],
    ["MdSms", "white"], ["MdComputer", "white"], ["MdWarning", "red"], ["MdSensors", "navy"], ["MdCellTower", "navy"],
    ["MdCellTower", "violet"], ["MdOilBarrel", "white"], ["MdCheckCircle", "white"],
  ];
  for (const [n, c] of need) ICONS[n + ":" + c] = await iconData(n, HEX[c]);
  const ic = (n, c = "white") => ICONS[n + ":" + c];

  // ---------- helpers (fresh option objects every call) ----------
  const nbs = (t) => t.replace(/(\d) (\d{3})/g, "$1\u00A0$2").replace(/(\d) (\d{3})/g, "$1\u00A0$2")
    .replace(/(\d) (h|L|mm|KB|MB|°C|dias)\b/g, "$1\u00A0$2").replace(/US\$ /g, "US$\u00A0").replace(/~(\d)/g, "~$1");
  const nb = (x) => typeof x === "string" ? nbs(x) : x.map((r) => Object.assign({}, r, { text: nbs(r.text) }));
  const T = (slide, text, o) => slide.addText(nb(text), Object.assign({ isTextBox: true, margin: 0, fontSize: 14, color: C.text1, valign: "top" }, o));
  const box = (slide, o) => slide.addShape(o.round === false ? S.rect : S.roundRect, Object.assign({}, o.round === false ? {} : { rectRadius: o.radius || 0.08 }, {
    x: o.x, y: o.y, w: o.w, h: o.h,
    fill: o.fill === null ? { type: "none" } : { color: o.fill || C.background2, transparency: o.ft || 0 },
    line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
    objectName: o.name,
  }));
  const tbox = (slide, text, o) => slide.addText(nb(text), Object.assign({
    isTextBox: true, shape: o.round === false ? S.rect : S.roundRect, rectRadius: o.radius || 0.08,
    fill: o.fill === null ? { type: "none" } : { color: o.fill || C.background2, transparency: o.ft || 0 },
    line: o.line ? { color: o.line, width: o.lw || 1, dashType: o.dash || "solid" } : { type: "none" },
    fontSize: 14, color: C.text1, align: "center", valign: "middle", margin: 4,
  }, o.opts || {}, { x: o.x, y: o.y, w: o.w, h: o.h, objectName: o.name }));
  const arrow = (slide, x1, y1, x2, y2, o = {}) => slide.addShape(S.line, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.max(Math.abs(x2 - x1), 0.001), h: Math.max(Math.abs(y2 - y1), 0.001),
    flipH: x2 < x1, flipV: y2 < y1,
    line: { color: o.color || C.accent1, width: o.width || 2, dashType: o.dash || "solid", endArrowType: o.head === false ? undefined : "triangle" },
    objectName: o.name,
  });
  const poly = (slide, pts, color, width, dash) => { for (let i = 1; i < pts.length; i++) arrow(slide, pts[i - 1][0], pts[i - 1][1], pts[i][0], pts[i][1], { color, width, dash, head: false }); };
  const circleIcon = (slide, x, y, d, iconName, fill, iconColor = "white") => {
    slide.addShape(S.ellipse, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" } });
    const p = d * 0.22;
    slide.addImage({ data: ic(iconName, iconColor), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };
  const content = (eyebrow, title, section) => {
    const s = pres.addSlide({ masterName: "VIGIA_CONTENT", sectionTitle: section });
    s.addText(eyebrow, { placeholder: "eyebrow" });
    s.addText(title, { placeholder: "title" });
    return s;
  };

  // =====================================================================
  // 1 · COVER
  // =====================================================================
  let SEC = "Abertura";
  pres.addSection({ title: SEC });
  let s = pres.addSlide({ masterName: "VIGIA_DARK", sectionTitle: SEC });
  s.addText("APRESENTAÇÃO TÉCNICA · OUTUBRO 2026", { placeholder: "eyebrow" });
  s.addText("", { placeholder: "title" });
  T(s, "PREPARADO PARA", { x: 8.2, y: 0.5, w: 2.35, h: 0.32, fontSize: 11, bold: true, fontFace: MONO, color: "C9CFDD", align: "right", charSpacing: 2 });
  s.addImage({ data: LOGO_W, x: 10.75, y: 0.38, w: 0.56 * LOGO_RATIO, h: 0.56, altText: "Sonangol", objectName: "logo-sonangol" });
  T(s, "VIGIA", { x: 0.6, y: 1.05, w: 8, h: 1.3, fontSize: 88, bold: true, fontFace: MONO, color: C.background1, charSpacing: 6 });
  T(s, "Telemetria remota de tanques para a rede Sonangol Distribuidora", { x: 0.6, y: 2.55, w: 9.2, h: 1.3, fontSize: 36, bold: true, color: C.background1 });
  T(s, "Como funciona — do tanque ao telemóvel", { x: 0.6, y: 3.95, w: 9, h: 0.5, fontSize: 20, color: C.accent2 });
  // flow strip of the five steps
  const coverSteps = [["MdSensors", "Sonda"], ["MdRouter", "Gateway"], ["MdCellTower", "Rede 4G"], ["MdCloud", "Plataforma"], ["MdPhoneIphone", "Ecrãs"]];
  coverSteps.forEach(([n, label], i) => {
    const x = 0.6 + i * 1.75;
    circleIcon(s, x, 4.85, 0.85, n, i === 3 ? HEX.violet : HEX.violet);
    T(s, label, { x: x - 0.3, y: 5.78, w: 1.45, h: 0.35, fontSize: 14, color: C.background1, align: "center" });
    if (i < 4) arrow(s, x + 0.95, 5.275, x + 1.65, 5.275, { color: C.accent2, width: 2 });
  });
  T(s, [{ text: "Aires Augusto Miguêns", options: { bold: true, breakLine: true } }, { text: "Cientista de Dados · airesmiguens@gmail.com · +244 931 391 063" }],
    { x: 0.6, y: 6.45, w: 7.5, h: 0.65, fontSize: 14, color: C.background1 });
  T(s, "Documento independente · não afiliado nem aprovado pela Sonangol", { x: 8.3, y: 6.75, w: 4.45, h: 0.35, fontSize: 11, color: "C9CFDD", align: "right" });
  s.addNotes("Abertura. Esta apresentação responde à pergunta que a Sonangol mais coloca: como monitorizar remotamente qualquer posto. Vou mostrar a tecnologia camada a camada — sonda, gateway, rede, plataforma e ecrãs — com exemplos concretos.");

  // =====================================================================
  // 2 · A PERGUNTA
  // =====================================================================
  s = pres.addSlide({ masterName: "VIGIA_DARK", sectionTitle: SEC });
  s.addText("A PERGUNTA", { placeholder: "eyebrow" });
  s.addText("Como ver, a partir da sede, quanto combustível há em cada tanque de cada posto — agora?", { placeholder: "title" });
  const q = [
    ["MdSensors", "Medidor no tanque", "Já existe na maioria dos postos modernos", HEX.violet, false],
    ["MdCellTower", "Rede móvel no posto", "Unitel, Africell, Movicel — ou internet fixa", HEX.violet, false],
    ["MdLink", "Falta só ligar os dois", "a uma plataforma central: é isso o Vigia", HEX.amber, true],
  ];
  q.forEach(([n, t1, t2, col, hi], i) => {
    const x = 0.6 + i * 4.1;
    box(s, { x, y: 3.35, w: 3.85, h: 2.9, fill: hi ? C.accent2 : C.background1, ft: hi ? 0 : 92, line: hi ? null : "3A4A6B", radius: 0.12 });
    if (hi) {
      s.addShape(S.ellipse, { x: x + 0.35, y: 3.7, w: 0.95, h: 0.95, fill: { color: C.background1 }, line: { type: "none" } });
      s.addImage({ data: ic("MdLink", "navy"), x: x + 0.56, y: 3.91, w: 0.53, h: 0.53 });
    } else circleIcon(s, x + 0.35, 3.7, 0.95, n, col);
    T(s, t1, { x: x + 0.35, y: 4.9, w: 3.2, h: 0.5, fontSize: 22, bold: true, color: hi ? C.text1 : C.background1 });
    T(s, t2, { x: x + 0.35, y: 5.4, w: 3.2, h: 0.7, fontSize: 16, color: hi ? C.text1 : "C9CFDD" });
    if (i < 2) T(s, "+", { x: x + 3.85, y: 4.45, w: 0.25, h: 0.6, fontSize: 28, bold: true, color: C.accent2, align: "center" });
  });
  s.addNotes("A boa notícia: não é preciso inventar nada. O medidor já está dentro do tanque na maioria dos postos modernos, e a rede móvel já chega ao posto. O que falta é ligar os dois a uma plataforma central — é exatamente isso que o Vigia faz.");

  // =====================================================================
  // 3 · JÁ EXISTE
  // =====================================================================
  SEC = "Já existe";
  pres.addSection({ title: SEC });
  s = content("01 · ISTO JÁ FOI FEITO?", "Sim — é tecnologia comercial, já usada em África", SEC);
  const vendors = [
    ["MdPublic", "Insite360", "Gilbarco Veeder-Root", "Global · Brasil · Europa", "Nuvem que lê os medidores de tanque de toda a rede de postos"],
    ["MdCellTower", "AFTECH", "Medição + 4G", "7 países africanos", "Envio por LTE, alertas de nível baixo, fuga ou furto"],
    ["MdHub", "AdvaIoT", "Gestão de pista", "África", "Tanques, bombas e caixa numa só plataforma"],
    ["MdStorage", "OTS Group", "Monitorização remota", "Reino Unido", "Mais de 300 tanques lidos à distância"],
  ];
  vendors.forEach(([n, name, sub, where, what], i) => {
    const x = 0.6 + i * 3.1;
    box(s, { x, y: 1.75, w: 2.85, h: 3.15, fill: C.background2, radius: 0.12 });
    circleIcon(s, x + 0.3, 2.05, 0.8, n, HEX.violet);
    T(s, name, { x: x + 0.3, y: 3.0, w: 2.35, h: 0.45, fontSize: 22, bold: true });
    T(s, sub, { x: x + 0.3, y: 3.42, w: 2.35, h: 0.3, fontSize: 13, color: C.text2 });
    T(s, where.toUpperCase(), { x: x + 0.3, y: 3.78, w: 2.4, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.accent1 });
    T(s, what, { x: x + 0.3, y: 4.1, w: 2.35, h: 0.75, fontSize: 14, color: C.text1 });
  });
  // AFTECH countries strip
  T(s, "AFTECH em: Senegal · Guiné · Serra Leoa · Libéria · Côte d’Ivoire · Zâmbia · África do Sul", { x: 0.6, y: 5.12, w: 12.1, h: 0.35, fontSize: 14, color: C.text2 });
  // takeaway band
  box(s, { x: 0.6, y: 5.65, w: 12.15, h: 1.05, fill: C.accent1, ft: 88, radius: 0.1 });
  T(s, "0", { x: 0.85, y: 5.7, w: 0.8, h: 0.95, fontSize: 48, bold: true, color: C.accent1, valign: "middle" });
  T(s, [{ text: "casos públicos numa rede inteira de distribuidor em Angola", options: { bold: true, breakLine: true } }, { text: "Resultados acima são afirmações dos fornecedores — o piloto vai medi-los na rede da Sonangol." }],
    { x: 1.65, y: 5.78, w: 10.9, h: 0.85, fontSize: 15, valign: "middle" });
  s.addNotes("A monitorização remota de tanques é um mercado estabelecido. O maior fabricante de medidores (Gilbarco Veeder-Root) vende uma plataforma na nuvem, e há empresas africanas a instalar sistemas com envio por 4G em sete países. O que não encontrei foi um caso público numa rede inteira em Angola — e os números africanos são dos próprios fornecedores. Por isso propomos um piloto que mede os resultados na rede da Sonangol. Porquê o Vigia: as plataformas dos fabricantes funcionam melhor com uma só marca; o Vigia junta todas as fontes, fica sob controlo da Sonangol e é adaptado à logística angolana.");

  // =====================================================================
  // 4 · ARQUITETURA
  // =====================================================================
  SEC = "Arquitetura";
  pres.addSection({ title: SEC });
  s = content("02 · ARQUITETURA GERAL", "Do tanque ao telemóvel em cinco passos", SEC);
  const zones = [["NO POSTO", 0.45, 4.95], ["NA REDE", 5.5, 2.35], ["NO SERVIDOR", 7.95, 2.35], ["NA MÃO", 10.4, 2.4]];
  zones.forEach(([label, x, w]) => {
    box(s, { x, y: 1.65, w, h: 3.95, fill: C.background2, radius: 0.12 });
    T(s, label, { x: x + 0.2, y: 1.8, w: w - 0.4, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.text2, charSpacing: 2 });
  });
  const steps = [
    ["MdSensors", "1 · Sonda", "No tanque: mede nível, água e temperatura"],
    ["MdRouter", "2 · Gateway", "No escritório: lê a consola a cada 10 min"],
    ["MdCellTower", "3 · Rede 4G", "Envia a leitura cifrada (≈ 1 KB)"],
    ["MdCloud", "4 · Plataforma", "Litros, previsão de rutura, alertas"],
    ["MdPhoneIphone", "5 · Ecrãs", "Painel, app e SMS para quem decide"],
  ];
  const colX = [0.7, 3.15, 5.65, 8.1, 10.55];
  steps.forEach(([n, t1, t2], i) => {
    const x = colX[i];
    circleIcon(s, x + 0.5, 2.3, 1.15, n, i === 3 ? HEX.amber : HEX.violet, "white");
    T(s, t1, { x, y: 3.65, w: 2.15, h: 0.45, fontSize: 18, bold: true, align: "center" });
    T(s, t2, { x: x + 0.05, y: 4.12, w: 2.05, h: 1.0, fontSize: 14, color: C.text2, align: "center" });
    if (i < 4) arrow(s, x + 1.75, 2.875, colX[i + 1] + 0.42, 2.875, { color: C.accent1, width: 2.5, dash: i === 2 || i === 1 ? "dash" : "solid" });
  });
  const facts = [["a cada 10 min", "uma leitura por tanque"], ["segundos", "do tanque ao ecrã"], ["só leitura", "nunca controla bombas ou preços"]];
  facts.forEach(([a, b], i) => {
    const x = 0.6 + i * 4.1;
    box(s, { x, y: 5.85, w: 3.85, h: 0.85, fill: C.accent1, ft: 88, radius: 0.1 });
    T(s, [{ text: a + "  ", options: { bold: true, color: C.accent1, fontSize: 18 } }, { text: b, options: { color: C.text1, fontSize: 14 } }], { x: x + 0.25, y: 5.85, w: 3.5, h: 0.85, valign: "middle" });
  });
  s.addNotes("Esta é a arquitetura inteira numa imagem. Passo 1: a sonda dentro do tanque mede continuamente. Passo 2: um pequeno gateway no escritório lê a consola do medidor a cada 10 minutos — só leitura. Passo 3: envia a leitura pela rede móvel, cifrada. Passo 4: a plataforma converte em litros, deteta vendas e entregas e prevê a hora de rutura. Passo 5: o alerta aparece no painel de despacho e no telemóvel do gestor regional. As linhas tracejadas são a parte que viaja por rede pública, sempre cifrada.");

  // =====================================================================
  // 5 · SONDA
  // =====================================================================
  SEC = "As camadas em detalhe";
  pres.addSection({ title: SEC });
  s = content("03 · CAMADA 1 — MEDIÇÃO NO TANQUE", "A sonda: um íman a flutuar, medido ao milímetro", SEC);
  // ground + pavement
  box(s, { x: 0.6, y: 2.3, w: 7.6, h: 4.25, fill: HEX.soil, round: false });
  box(s, { x: 0.6, y: 2.15, w: 7.6, h: 0.15, fill: "C9CCCB", round: false });
  T(s, "pavimento da pista", { x: 0.7, y: 1.8, w: 2.2, h: 0.28, fontSize: 11, color: C.text2 });
  // zone 1 + manhole
  box(s, { x: 3.65, y: 1.7, w: 1.3, h: 1.62, fill: null, line: HEX.red, lw: 1.5, dash: "dash", round: false });
  T(s, "ZONA 1", { x: 5.05, y: 1.72, w: 1.0, h: 0.28, fontSize: 11, bold: true, color: C.accent3 });
  box(s, { x: 3.9, y: 2.15, w: 0.8, h: 0.98, fill: "D8DBD8", line: HEX.navy, lw: 1, round: false });
  box(s, { x: 3.8, y: 2.05, w: 1.0, h: 0.1, fill: C.text1, round: false });
  // tank
  box(s, { x: 1.0, y: 3.1, w: 6.6, h: 3.25, fill: C.background1, line: HEX.navy, lw: 2.5, radius: 0.35 });
  box(s, { x: 1.06, y: 4.3, w: 6.48, h: 1.98, fill: C.accent2, radius: 0.28 });
  box(s, { x: 1.06, y: 5.98, w: 6.48, h: 0.3, fill: C.accent5, round: false });
  T(s, "vapor", { x: 1.35, y: 3.35, w: 1.2, h: 0.3, fontSize: 12, color: C.text2 });
  T(s, "COMBUSTÍVEL", { x: 1.35, y: 4.95, w: 2.2, h: 0.4, fontSize: 16, bold: true });
  T(s, "ÁGUA", { x: 1.35, y: 5.99, w: 1.2, h: 0.28, fontSize: 11, bold: true, color: C.background1, valign: "middle" });
  T(s, "ZONA 0 · interior do tanque", { x: 5.05, y: 3.35, w: 2.4, h: 0.3, fontSize: 11, bold: true, color: C.accent3 });
  // probe
  s.addShape(S.line, { x: 4.3, y: 2.2, w: 0.001, h: 4.05, line: { color: HEX.navy, width: 4 } });
  box(s, { x: 4.22, y: 2.0, w: 0.16, h: 0.2, fill: C.text1, round: false });
  box(s, { x: 4.05, y: 4.15, w: 0.5, h: 0.3, fill: C.background1, line: HEX.navy, lw: 2, radius: 0.12 });
  box(s, { x: 4.1, y: 5.93, w: 0.4, h: 0.22, fill: "DDEBFA", line: HEX.navy, lw: 2, radius: 0.1 });
  [3.6, 4.9, 5.45].forEach((y) => s.addShape(S.ellipse, { x: 4.24, y, w: 0.12, h: 0.12, fill: { color: C.accent3 }, line: { type: "none" } }));
  // office + cable
  tbox(s, "Consola ATG · escritório", { x: 6.15, y: 1.6, w: 2.0, h: 0.42, fill: C.background1, line: HEX.navy, lw: 1, opts: { fontSize: 11, bold: true } });
  arrow(s, 4.38, 2.0, 5.85, 2.0, { color: C.text1, width: 1.25, dash: "dash", head: false });
  arrow(s, 5.85, 2.0, 5.85, 1.81, { color: C.text1, width: 1.25, dash: "dash", head: false });
  arrow(s, 5.85, 1.81, 6.15, 1.81, { color: C.text1, width: 1.25, dash: "dash" });
  // letter badges
  const badge = (slide, L, x, y) => tbox(slide, L, { x, y, w: 0.34, h: 0.34, fill: C.accent1, opts: { shape: S.ellipse, fontSize: 13, bold: true, color: C.background1, margin: 0 } });
  badge(s, "A", 4.68, 4.13); badge(s, "B", 4.62, 5.86); badge(s, "C", 3.82, 4.78); badge(s, "D", 3.82, 2.55);
  // legend
  const legend = [
    ["A", "Flutuador de produto", "Segue a superfície do combustível"],
    ["B", "Flutuador de água", "Mais denso: só flutua na água do fundo"],
    ["C", "Sensores de temperatura", "Vários pontos ao longo da haste"],
    ["D", "Haste magnetostritiva", "Pulso elétrico + íman = posição ao milímetro"],
  ];
  legend.forEach(([L, t1, t2], i) => {
    const y = 1.75 + i * 0.98;
    badge(s, L, 8.6, y + 0.04);
    T(s, t1, { x: 9.1, y, w: 3.65, h: 0.36, fontSize: 16, bold: true });
    T(s, t2, { x: 9.1, y: y + 0.37, w: 3.65, h: 0.5, fontSize: 14, color: C.text2 });
  });
  tbox(s, [{ text: "A consola converte", options: { bold: true, breakLine: true } }, { text: "altura → litros (tabela de arqueação) → volume a 15 °C" }],
    { x: 8.6, y: 5.75, w: 4.15, h: 0.95, fill: C.accent1, opts: { color: C.background1, fontSize: 14, align: "left", margin: 10 } });
  T(s, "Corte esquemático, fora de escala. Zonas 0 e 1: atmosfera explosiva — todo o equipamento aí é certificado ATEX/IECEx.", { x: 0.6, y: 6.6, w: 7.6, h: 0.3, fontSize: 10, color: C.text2 });
  s.addNotes("Os tanques são subterrâneos, tipicamente 20 a 50 mil litros, um por produto. A sonda é uma haste vertical instalada pela boca de visita. Tem dois flutuadores com ímanes: um segue a superfície do combustível, o outro, mais denso, só flutua na água do fundo. A consola envia um pulso elétrico pela haste; quando passa pelo íman cria uma onda que volta ao topo, e o tempo de regresso dá a altura com precisão da ordem do milímetro. Com a tabela de arqueação, a consola converte altura em litros e corrige para 15 °C, porque o combustível dilata com o calor. O interior do tanque é Zona 0 e a boca de visita é Zona 1: atmosfera explosiva, por isso tudo o que lá estiver é certificado.");

  // =====================================================================
  // 6 · OPÇÕES DE MEDIÇÃO
  // =====================================================================
  s = content("03 · CAMADA 1 — OPÇÕES", "Cinco formas de medir: escolhe-se posto a posto", SEC);
  const hdrY = 1.7;
  [["TECNOLOGIA", 0.85], ["PRECISÃO", 5.35], ["NO TANQUE", 8.05], ["QUANDO USAR", 10.25]].forEach(([h, x]) =>
    T(s, h, { x, y: hdrY, w: 2.4, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.text2, charSpacing: 2 }));
  const opts = [
    ["Medidor ATG já existente", "Sonda magnetostritiva + consola", 5, "Nada a instalar", "Postos modernos e requalificados", true],
    ["Sonda magnetostritiva nova", "Mesma tecnologia, instalada de raiz", 5, "Sonda + consola", "Postos sem ATG", false],
    ["Radar sem contacto", "Micro-ondas a partir do topo", 4, "Radar no topo", "Sem acesso para a haste", false],
    ["Ultrassónico", "Som a partir do topo", 3, "Sensor no topo", "Tanques de superfície (gasóleo)", false],
    ["Vareta manual + app", "Funcionário regista 2–3 vezes/dia", 1, "Nada", "Passo intermédio, já amanhã", false],
  ];
  opts.forEach(([name, sub, prec, hw, when, ideal], i) => {
    const y = 2.1 + i * 0.92;
    box(s, { x: 0.6, y, w: 12.15, h: 0.8, fill: ideal ? C.accent1 : C.background2, ft: ideal ? 86 : 0, radius: 0.1 });
    T(s, name, { x: 0.85, y: y + 0.08, w: 4.3, h: 0.35, fontSize: 16, bold: true });
    T(s, sub, { x: 0.85, y: y + 0.42, w: 4.3, h: 0.3, fontSize: 13, color: C.text2 });
    for (let k = 0; k < 5; k++) box(s, { x: 5.35 + k * 0.42, y: y + 0.29, w: 0.34, h: 0.22, fill: k < prec ? C.accent1 : "D5DAE1", radius: 0.05 });
    tbox(s, hw, { x: 8.05, y: y + 0.2, w: 1.95, h: 0.4, fill: hw.startsWith("Nada") ? C.accent4 : C.background1, line: hw.startsWith("Nada") ? null : HEX.hair, opts: { fontSize: 13, bold: true, color: hw.startsWith("Nada") ? C.background1 : C.text1 } });
    T(s, when, { x: 10.25, y: y + 0.1, w: 2.4, h: 0.6, fontSize: 14, valign: "middle" });
    if (ideal) tbox(s, "IDEAL", { x: 4.15, y: y + 0.1, w: 0.9, h: 0.3, fill: C.accent1, opts: { fontSize: 11, bold: true, color: C.background1, fontFace: MONO } });
  });
  s.addNotes("O piloto começa por um levantamento: em cada posto vemos o que já existe. O caso ideal é o posto que já tem medidor automático (ATG): não se instala nada no tanque, só se liga um cabo de dados à consola. Onde não há ATG, instala-se uma sonda nova ou um radar no topo, sempre certificados. O ultrassónico é pouco adequado a gasolina por causa dos vapores. E a vareta manual com registo numa app é um passo intermédio que pode começar já, sem hardware.");

  // =====================================================================
  // 7 · DENTRO DO POSTO
  // =====================================================================
  s = content("04 · CAMADA 2 — DENTRO DO POSTO", "Um único equipamento novo, numa rede separada", SEC);
  // forecourt
  box(s, { x: 0.6, y: 1.6, w: 5.6, h: 2.3, fill: C.background2, radius: 0.12 });
  T(s, "PISTA · zona classificada", { x: 0.85, y: 1.75, w: 4, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.text2 });
  [0.95, 2.75, 4.55].forEach((x, i) => {
    box(s, { x, y: 2.2, w: 1.3, h: 1.45, fill: C.background1, line: HEX.hair, radius: 0.1 });
    s.addImage({ data: ic("MdLocalGasStation", "navy"), x: x + 0.38, y: 2.32, w: 0.55, h: 0.55 });
    T(s, "Bomba " + (i + 1), { x, y: 3.05, w: 1.3, h: 0.35, fontSize: 13, align: "center" });
  });
  // tanks
  box(s, { x: 0.6, y: 4.1, w: 5.6, h: 2.6, fill: C.accent2, ft: 85, radius: 0.12 });
  T(s, "TANQUES SUBTERRÂNEOS", { x: 0.85, y: 4.25, w: 4, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.accent6 });
  [["T1 Gasolina", 0.85], ["T2 Gasóleo", 2.65], ["T3 Gasóleo", 4.45]].forEach(([n, x]) => {
    box(s, { x, y: 4.75, w: 1.55, h: 0.85, fill: C.accent2, line: HEX.navy, lw: 1.25, radius: 0.3 });
    s.addShape(S.line, { x: x + 1.15, y: 4.62, w: 0.001, h: 0.9, line: { color: HEX.navy, width: 2.5 } });
    T(s, n, { x, y: 5.7, w: 1.55, h: 0.3, fontSize: 13, align: "center" });
  });
  T(s, "uma sonda por tanque", { x: 0.85, y: 6.2, w: 3, h: 0.3, fontSize: 12, color: C.text2 });
  // office
  box(s, { x: 6.6, y: 1.6, w: 6.15, h: 5.1, fill: C.background1, line: HEX.navy, lw: 1.25, radius: 0.12 });
  T(s, "ESCRITÓRIO · zona segura", { x: 6.85, y: 1.72, w: 4, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.text2 });
  box(s, { x: 6.85, y: 2.1, w: 5.65, h: 2.55, fill: C.accent1, ft: 90, line: HEX.violet, lw: 1.25, dash: "dash", radius: 0.1 });
  T(s, "REDE DE MONITORIZAÇÃO · isolada", { x: 7.05, y: 2.2, w: 4.5, h: 0.3, fontSize: 11, bold: true, fontFace: MONO, color: C.accent1 });
  tbox(s, [{ text: "Consola ATG", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "ex.: TLS-350 / 450", options: { fontSize: 12, color: C.text2 } }],
    { x: 7.1, y: 2.75, w: 2.1, h: 1.0, fill: C.background1, line: HEX.navy, lw: 1.25 });
  tbox(s, [{ text: "Gateway 4G", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "dual-SIM · TLS", options: { fontSize: 12, color: "C9CFDD" } }],
    { x: 10.1, y: 2.75, w: 2.1, h: 1.0, fill: C.text1, opts: { color: C.background1 } });
  arrow(s, 9.2, 3.25, 10.1, 3.25, { color: C.accent1, width: 2.5 });
  T(s, "RS-232", { x: 9.2, y: 2.88, w: 0.9, h: 0.28, fontSize: 11, bold: true, color: C.accent1, align: "center" });
  T(s, "só leitura", { x: 9.2, y: 3.33, w: 0.9, h: 0.28, fontSize: 10, color: C.accent1, align: "center" });
  tbox(s, "UPS · bateria 8–24 h", { x: 10.1, y: 3.95, w: 2.1, h: 0.5, fill: C.background1, line: HEX.hair, opts: { fontSize: 12 } });
  s.addImage({ data: ic("MdCellTower", "violet"), x: 11.95, y: 2.15, w: 0.4, h: 0.4 });
  arrow(s, 12.15, 2.58, 12.15, 2.75, { color: C.accent1, width: 1.5, head: false });
  T(s, "antena", { x: 11.2, y: 2.25, w: 0.7, h: 0.25, fontSize: 10, color: C.accent1, align: "right" });
  // probe cables → console
  arrow(s, 6.2, 4.9, 8.15, 4.9, { color: C.text1, width: 1.5, dash: "dash", head: false });
  arrow(s, 8.15, 4.9, 8.15, 3.75, { color: C.text1, width: 1.5, dash: "dash" });
  T(s, "cabos das sondas", { x: 8.3, y: 4.75, w: 1.7, h: 0.28, fontSize: 10, color: C.text2 });
  tbox(s, "SEM LIGAÇÃO", { x: 10.25, y: 4.78, w: 1.8, h: 0.3, fill: C.accent3, opts: { fontSize: 11, bold: true, color: C.background1 } });
  // shop network
  box(s, { x: 6.85, y: 5.2, w: 5.65, h: 1.35, fill: C.background2, line: HEX.hair, radius: 0.1 });
  T(s, "REDE DA LOJA · o Vigia não lhe toca", { x: 7.05, y: 5.3, w: 5, h: 0.28, fontSize: 11, bold: true, fontFace: MONO, color: C.text2 });
  [["POS / caixa", 7.05, 1.65], ["Controlador de bombas", 8.85, 1.8], ["Router da loja", 10.8, 1.5]].forEach(([n, x, w]) =>
    tbox(s, n, { x, y: 5.7, w, h: 0.65, fill: C.background1, line: HEX.hair, opts: { fontSize: 12, color: C.text2 } }));
  s.addNotes("Dentro do posto, o Vigia acrescenta um só equipamento: o gateway, no escritório, ao lado da consola do medidor. Liga-se à consola por um cabo série (RS-232) e só pede leituras. Fica numa rede própria, separada da rede da loja: não toca no sistema de caixa, no controlador das bombas nem no router da loja. Tem antena externa e bateria para os cortes de energia. Numa fase 2 opcional, e só com aprovação da TI da Sonangol, pode ler as vendas do controlador de bombas para a reconciliação.");

  // =====================================================================
  // 8 · GATEWAY
  // =====================================================================
  s = content("04 · CAMADA 2 — O GATEWAY", "O gateway: pequeno, robusto e só de leitura", SEC);
  // device illustration
  arrow(s, 1.4, 1.85, 1.4, 2.35, { color: C.text1, width: 4, head: false });
  arrow(s, 4.4, 1.85, 4.4, 2.35, { color: C.text1, width: 4, head: false });
  box(s, { x: 0.6, y: 2.3, w: 4.6, h: 2.7, fill: C.text1, radius: 0.18 });
  T(s, "GATEWAY 4G", { x: 0.95, y: 2.55, w: 3.0, h: 0.4, fontSize: 18, bold: true, fontFace: MONO, color: C.background1 });
  [HEX.green, HEX.amber, HEX.lilac].forEach((c, i) => s.addShape(S.ellipse, { x: 4.2 + i * 0.25, y: 2.65, w: 0.15, h: 0.15, fill: { color: c }, line: { type: "none" } }));
  [["SIM 1", 0.95], ["SIM 2", 1.8], ["RS-232", 2.65], ["ETH", 3.6], ["DC", 4.3]].forEach(([p, x]) =>
    tbox(s, p, { x, y: 4.2, w: p === "RS-232" ? 0.85 : p === "DC" ? 0.6 : 0.75, h: 0.45, fill: "2A3A5E", line: "8F98AD", lw: 0.75, radius: 0.05, opts: { fontSize: 11, bold: true, color: C.background1, fontFace: MONO } }));
  T(s, "Do tamanho de um router · montado no escritório, ao lado da consola do medidor", { x: 0.6, y: 5.2, w: 4.6, h: 0.7, fontSize: 14, color: C.text2 });
  const specs = [
    ["MdSignalCellularAlt", "4G dual-SIM", "Dois operadores, troca automática"],
    ["MdSettingsInputComponent", "RS-232 · Ethernet", "Liga-se à consola do medidor"],
    ["MdBatteryChargingFull", "Bateria", "Continua a medir nos cortes"],
    ["MdStorage", "30 dias de memória", "Nada se perde se a rede cair"],
    ["MdLock", "TLS + certificado", "Cifrado, com identidade única"],
    ["MdPrecisionManufacturing", "Gama industrial", "−20 a +70 °C · calha DIN"],
  ];
  specs.forEach(([n, t1, t2], i) => {
    const x = 5.7 + (i % 3) * 2.375, y = 1.75 + Math.floor(i / 3) * 2.3;
    box(s, { x, y, w: 2.2, h: 2.1, fill: C.background2, radius: 0.12 });
    circleIcon(s, x + 0.25, y + 0.25, 0.65, n, HEX.violet);
    T(s, t1, { x: x + 0.25, y: y + 1.0, w: 1.85, h: 0.4, fontSize: 16, bold: true });
    T(s, t2, { x: x + 0.25, y: y + 1.38, w: 1.85, h: 0.6, fontSize: 13, color: C.text2 });
  });
  T(s, "Exemplos de classe: Teltonika RUT, Moxa, Advantech — escolha final após o levantamento.", { x: 5.7, y: 6.4, w: 7.05, h: 0.3, fontSize: 11, color: C.text2 });
  s.addNotes("O gateway é do tamanho de um router. Tem dois cartões SIM de operadores diferentes e troca sozinho se um falhar. Liga-se à consola por porta série ou rede. Tem bateria, guarda 30 dias de leituras se a rede cair, comunica cifrado com um certificado único, e é de gama industrial para calor e poeira. Existem várias marcas desta classe; a escolha final faz-se depois do levantamento.");

  // =====================================================================
  // 9 · REDE NACIONAL
  // =====================================================================
  s = content("05 · CAMADA 3 — REDE NACIONAL", "Topologia em estrela: cada posto fala direto com a plataforma", SEC);
  const stations = [
    ["Talatona · Luanda", "Unitel + Africell (reserva)"], ["Kilamba Kiaxi · Luanda", "Unitel + Movicel (reserva)"],
    ["Viana · Luanda", "Africell + fibra do posto"], ["Lobito · Benguela", "Movicel + Unitel (reserva)"], ["Posto rural (EN)", "VSAT · sem cobertura móvel"],
  ];
  const stY = [1.65, 2.5, 3.35, 4.2, 5.05];
  stations.forEach(([a, b], i) => tbox(s, [{ text: a, options: { bold: true, fontSize: 13, breakLine: true } }, { text: b, options: { fontSize: 12, color: C.text2 } }],
    { x: 0.6, y: stY[i], w: 2.9, h: 0.65, fill: C.background1, line: HEX.navy, lw: 1, opts: { align: "left", margin: [3, 8, 3, 8] } }));
  T(s, "… até 326 postos", { x: 0.6, y: 5.85, w: 2.9, h: 0.3, fontSize: 12, italic: true, color: C.text2 });
  const ops = [["Unitel", 1.85], ["Africell", 2.9], ["Movicel", 3.95], ["Satélite", 5.0]];
  ops.forEach(([n, y]) => tbox(s, n, { x: 4.3, y, w: 2.0, h: 0.6, fill: C.accent1, ft: 88, line: HEX.violet, lw: 1, radius: 0.3, opts: { fontSize: 15, bold: true } }));
  arrow(s, 3.5, 1.97, 4.3, 2.08, { width: 1.5 });
  arrow(s, 3.5, 2.82, 4.3, 2.22, { width: 1.5 });
  arrow(s, 3.5, 3.67, 4.3, 3.2, { width: 1.5 });
  arrow(s, 3.5, 4.52, 4.3, 4.25, { width: 1.5 });
  arrow(s, 3.5, 5.37, 4.3, 5.3, { width: 1.5 });
  arrow(s, 3.5, 2.05, 4.3, 3.05, { width: 1.25, dash: "dash", color: C.text2 });
  tbox(s, [{ text: "APN privada", options: { bold: true, fontSize: 16, breakLine: true } }, { text: "ou Internet", options: { fontSize: 13, breakLine: true } }, { text: "túnel cifrado TLS", options: { fontSize: 13, color: C.accent1 } }],
    { x: 7.0, y: 2.75, w: 2.0, h: 1.5, fill: C.background1, line: HEX.navy, lw: 1.25, radius: 0.2 });
  arrow(s, 6.3, 2.15, 7.0, 3.0, { width: 1.5 });
  arrow(s, 6.3, 3.2, 7.0, 3.35, { width: 1.5 });
  arrow(s, 6.3, 4.25, 7.0, 3.7, { width: 1.5 });
  arrow(s, 6.3, 5.3, 7.0, 4.05, { width: 1.5 });
  box(s, { x: 9.35, y: 3.0, w: 0.9, h: 1.0, fill: C.accent3, radius: 0.08 });
  s.addImage({ data: ic("MdLock", "white"), x: 9.6, y: 3.12, w: 0.4, h: 0.4 });
  T(s, "firewall", { x: 9.35, y: 3.6, w: 0.9, h: 0.3, fontSize: 11, bold: true, color: C.background1, align: "center" });
  arrow(s, 9.0, 3.5, 9.35, 3.5, { width: 2 });
  arrow(s, 10.25, 3.5, 10.6, 3.5, { width: 2 });
  box(s, { x: 10.6, y: 1.65, w: 2.15, h: 4.2, fill: C.accent1, ft: 88, line: HEX.violet, lw: 1.25, radius: 0.12 });
  T(s, "PLATAFORMA VIGIA", { x: 10.75, y: 1.78, w: 1.9, h: 0.3, fontSize: 12, bold: true, color: C.accent1, fontFace: MONO });
  ["Broker MQTT · 8883", "Ingestão", "Séries temporais", "Previsão", "Gestão de gateways", "API · Web · App · SMS"].forEach((n, i) =>
    tbox(s, n, { x: 10.75, y: 2.2 + i * 0.6, w: 1.85, h: 0.48, fill: i === 5 ? C.accent1 : C.background1, line: i === 5 ? null : HEX.lilac, opts: { fontSize: 12, bold: i === 5, color: i === 5 ? C.background1 : C.text1 } }));
  tbox(s, "Sistemas Sonangol · ERP / logística (fase 2)", { x: 10.6, y: 6.05, w: 2.15, h: 0.62, fill: null, line: HEX.navy, lw: 1, dash: "dash", opts: { fontSize: 11 } });
  arrow(s, 11.675, 5.85, 11.675, 6.05, { width: 1.25, dash: "dash", color: C.text1, head: false });
  T(s, "Linha contínua: ligação principal · tracejada: reserva automática", { x: 4.3, y: 6.35, w: 6, h: 0.3, fontSize: 11, color: C.text2 });
  s.addNotes("Cada posto fala diretamente com a plataforma central — uma topologia em estrela. Se um posto perder a rede, os outros continuam normalmente. Cada posto tem dois caminhos: dois operadores, ou fibra mais 4G, com troca automática. Postos rurais sem cobertura usam satélite. Com uma APN privada contratada ao operador, os gateways nem sequer ficam visíveis na Internet pública. Tudo entra por um firewall na plataforma, que pode estar no datacenter da Sonangol ou numa nuvem escolhida pela Sonangol.");

  // =====================================================================
  // 10 · PROTOCOLOS
  // =====================================================================
  s = content("05 · CAMADA 3 — PROTOCOLOS", "A conversa entre o posto e a plataforma, passo a passo", SEC);
  const code = [
    ["1 · O gateway pergunta", "RS-232 · protocolo Veeder-Root (padrão do setor)", "<SOH>i20100\n\nfunção 201 =\nrelatório de inventário\nde todos os tanques"],
    ["2 · A consola responde", "em menos de um segundo", "T1 GASOLINA\n VOL 18 420 L  ALT 1 236 mm\n ÁGUA 6 mm    28,4 °C\nT2 GASÓLEO\n VOL  5 110 L  ALT   498 mm\n ÁGUA 0 mm    29,1 °C"],
    ["3 · O gateway envia", "MQTT sobre TLS · porta 8883", "{ \"posto\": \"SD-LDA-0142\",\n  \"tanque\": \"T2\",\n  \"volume_15c_l\": 5062,\n  \"agua_mm\": 0,\n  \"temp_c\": 29.1,\n  \"sinal_dbm\": -79 }"],
  ];
  code.forEach(([t1, t2, c], i) => {
    const x = 0.6 + i * 4.15;
    T(s, t1, { x, y: 1.7, w: 3.85, h: 0.4, fontSize: 18, bold: true });
    T(s, t2, { x, y: 2.1, w: 3.85, h: 0.3, fontSize: 12, color: C.accent1 });
    tbox(s, c, { x, y: 2.5, w: 3.85, h: 2.45, fill: C.text1, radius: 0.1, opts: { fontFace: MONO, fontSize: 13, color: "E6EAF2", align: "left", valign: "top", margin: 12 } });
    if (i < 2) arrow(s, x + 3.88, 3.72, x + 4.12, 3.72, { width: 2.5 });
  });
  const stats = [["10 min", "entre leituras"], ["≈ 1 KB", "por leitura"], ["≈ 5 MB", "de dados por mês, por posto"]];
  stats.forEach(([a, b], i) => {
    const x = 0.6 + i * 4.15;
    box(s, { x, y: 5.3, w: 3.85, h: 1.35, fill: C.background2, radius: 0.12 });
    T(s, a, { x: x + 0.3, y: 5.38, w: 3.3, h: 0.75, fontSize: 40, bold: true, color: C.accent1 });
    T(s, b, { x: x + 0.3, y: 6.1, w: 3.3, h: 0.4, fontSize: 14, color: C.text2 });
  });
  s.addNotes("Isto é a conversa real, simplificada. O gateway envia à consola o comando i201 — o relatório de inventário do protocolo Veeder-Root, que é o padrão do setor e é imitado por muitas outras marcas. A consola responde com volume, altura, água e temperatura de cada tanque. O gateway transforma isso numa mensagem pequena e envia-a à plataforma por MQTT, um protocolo leve feito para redes móveis instáveis, cifrado com TLS. São cerca de 1 KB por leitura: com três tanques a cada 10 minutos, dá uns 5 MB por mês por posto — um pacote de dados M2M básico chega.");

  // =====================================================================
  // 11 · PROCESSAMENTO
  // =====================================================================
  SEC = "O que a plataforma calcula";
  pres.addSection({ title: SEC });
  s = content("06 · CAMADA 4 — PROCESSAMENTO", "De uma leitura a uma decisão, em segundos", SEC);
  const chain = [["Leitura", "bruta"], ["Validação", "valores impossíveis"], ["Litros", "a 15 °C"], ["Eventos", "venda · entrega"], ["Previsão", "hora de rutura"], ["Prioridade", "tempo de viagem"], ["Alerta", "ação recomendada"]];
  chain.forEach(([a, b], i) => {
    const x = 0.6 + i * 1.775, last = i === 6;
    tbox(s, [{ text: a, options: { bold: true, fontSize: 15, breakLine: true } }, { text: b, options: { fontSize: 12, color: last ? "E6E2FA" : C.text2 } }],
      { x, y: 1.75, w: 1.5, h: 1.05, fill: last ? C.accent1 : C.background2, opts: { color: last ? C.background1 : C.text1 } });
    if (i < 6) arrow(s, x + 1.52, 2.275, x + 1.755, 2.275, { width: 2 });
  });
  const ev = [
    ["MdTrendingDown", "Venda", "O nível desce aos poucos, ao ritmo das bombas", HEX.navy, [100, 96, 91, 85, 80, 74, 69, 63, 58, 52]],
    ["MdTrendingUp", "Entrega", "O nível sobe de repente — e confere-se com a fatura", HEX.green, [40, 37, 34, 31, 90, 88, 85, 82, 79, 76]],
    ["MdReportProblem", "Anomalia", "O nível desce sem vendas (ex.: de madrugada)", HEX.red, [80, 80, 80, 79, 74, 69, 64, 64, 64, 64]],
  ];
  ev.forEach(([n, t1, t2, col, data], i) => {
    const x = 0.6 + i * 4.15;
    box(s, { x, y: 3.15, w: 3.85, h: 3.55, fill: C.background1, line: HEX.hair, lw: 1, radius: 0.12 });
    circleIcon(s, x + 0.25, 3.35, 0.6, n, col);
    T(s, t1, { x: x + 1.0, y: 3.45, w: 2.6, h: 0.4, fontSize: 18, bold: true });
    arrow(s, x + 0.25, 5.62, x + 3.6, 5.62, { color: "D5DAE1", width: 1, head: false });
    poly(s, data.map((v, k) => [x + 0.25 + k * (3.35 / 9), 5.6 - (v / 100) * 1.45]), col, 3.5);
    T(s, t2, { x: x + 0.25, y: 5.8, w: 3.4, h: 0.7, fontSize: 14, color: C.text2 });
  });
  s.addNotes("Cada leitura passa por esta cadeia: valida-se, converte-se em litros a 15 °C, e compara-se com as anteriores para reconhecer eventos. Uma venda é uma descida gradual. Uma entrega é uma subida brusca — e a plataforma confere o volume recebido com a fatura. Uma anomalia é uma descida sem vendas, por exemplo de madrugada, o que pode indicar fuga ou desvio. Depois, a previsão calcula a hora de rutura e a prioridade junta o tempo de viagem do camião.");

  // =====================================================================
  // 12 · GRÁFICO
  // =====================================================================
  s = content("06 · CAMADA 4 — EXEMPLO", "Um tanque, 30 horas: a plataforma diz quando despachar", SEC);
  const labels = [], hist = [], fc = [], crit = [];
  const H = [12000, 11950, 11900, 11800, 11650, 11400, 11000, 10200, 9000, 7400, 5800, 24700, 23800, 22500, 21000, 19400, 17800, 16000, 14000, 12100];
  const F = { 19: 12100, 20: 10800, 21: 9600, 22: 8300, 23: 7000, 24: 5800, 25: 4700, 26: 3900, 27: 3300, 28: 2800, 29: 2400 };
  for (let h = 0; h < 30; h++) {
    const hh = String(h % 24).padStart(2, "0") + "h";
    labels.push(h === 0 ? "ter 00h" : h === 24 ? "qua 00h" : hh);
    hist.push(h < H.length ? H[h] : null);
    fc.push(F[h] !== undefined ? F[h] : null);
    crit.push(4500);
  }
  {
    const px0 = 1.5, pw = 7.4, py0 = 2.0, ph = 3.95;
    const X = (h) => px0 + (h * pw) / 29, Y = (L) => py0 + ph - (L / 30000) * ph;
    T(s, "Tanque T2 · gasóleo · 30 000 L (dados ilustrativos)", { x: 0.6, y: 1.55, w: 8.3, h: 0.35, fontSize: 14, bold: true });
    box(s, { x: px0, y: Y(4500), w: pw, h: Y(0) - Y(4500), fill: C.accent3, ft: 90, round: false });
    [0, 7500, 15000, 22500, 30000].forEach((L) => {
      arrow(s, px0, Y(L), px0 + pw, Y(L), { color: "E5E8EC", width: 0.75, head: false });
      T(s, { 0: "0 L", 7500: "7 500", 15000: "15 000", 22500: "22 500", 30000: "30 000" }[L], { x: 0.5, y: Y(L) - 0.14, w: 0.88, h: 0.28, fontSize: 11, color: C.text2, align: "right" });
    });
    arrow(s, px0, Y(4500), px0 + pw, Y(4500), { color: C.accent3, width: 1.5, dash: "dash", head: false });
    T(s, "nível crítico 15% (4 500 L)", { x: px0 + 0.1, y: Y(4500) + 0.06, w: 3, h: 0.28, fontSize: 11, bold: true, color: C.accent3 });
    [[0, "ter 00h"], [6, "06h"], [12, "12h"], [18, "18h"], [24, "qua 00h"], [29, "05h"]].forEach(([h, lab]) =>
      T(s, lab, { x: X(h) - 0.5, y: py0 + ph + 0.08, w: 1.0, h: 0.28, fontSize: 11, color: C.text2, align: "center" }));
    poly(s, H.map((L, h) => [X(h), Y(L)]), C.text1, 3);
    poly(s, Object.keys(F).map(Number).sort((p, q) => p - q).map((h) => [X(h), Y(F[h])]), C.accent1, 3, "dash");
    s.addShape(S.ellipse, { x: X(19) - 0.09, y: Y(12100) - 0.09, w: 0.18, h: 0.18, fill: { color: C.text1 }, line: { color: C.background1, width: 1.5 } });
    s.addShape(S.ellipse, { x: X(25.2) - 0.1, y: Y(4500) - 0.1, w: 0.2, h: 0.2, fill: { color: C.accent3 }, line: { color: C.background1, width: 1.5 } });
    T(s, "noite: consumo baixo", { x: X(0.3), y: Y(12000) - 0.42, w: 2.4, h: 0.28, fontSize: 11, color: C.text2 });
    T(s, "agora", { x: X(19) + 0.12, y: Y(12100) - 0.38, w: 0.8, h: 0.28, fontSize: 11, bold: true });
    T(s, "previsão", { x: X(21.5), y: Y(9000) - 0.45, w: 1.0, h: 0.28, fontSize: 11, bold: true, color: C.accent1 });
    T(s, "entrega", { x: X(11) + 0.1, y: Y(24700) - 0.32, w: 1.0, h: 0.28, fontSize: 11, bold: true, color: C.accent4 });
    // legend
    [["Nível medido", C.text1, "solid"], ["Previsão", C.accent1, "dash"], ["Nível crítico", C.accent3, "dash"]].forEach(([lab, col, d], i) => {
      const lx = 1.5 + i * 2.3;
      arrow(s, lx, 6.55, lx + 0.45, 6.55, { color: col, width: 3, dash: d, head: false });
      T(s, lab, { x: lx + 0.55, y: 6.41, w: 1.6, h: 0.28, fontSize: 12, color: C.text2 });
    });
  }
  const notes12 = [
    ["MdLocalShipping", HEX.green, "Entrega detetada · 10h05", "+19 700 L · fatura 20 000 L"],
    ["MdSensors", HEX.navy, "Agora · terça 19h", "12 100 L no tanque"],
    ["MdSchedule", HEX.red, "Crítico previsto · quarta 01h10", "Camião a 3 h do depósito"],
  ];
  notes12.forEach(([n, col, t1, t2], i) => {
    const y = 1.65 + i * 1.32;
    box(s, { x: 9.35, y, w: 3.4, h: 1.15, fill: C.background2, radius: 0.12 });
    circleIcon(s, 9.55, y + 0.25, 0.65, n, col);
    T(s, t1, { x: 10.35, y: y + 0.15, w: 2.3, h: 0.5, fontSize: 14, bold: true });
    T(s, t2, { x: 10.35, y: y + 0.62, w: 2.3, h: 0.45, fontSize: 13, color: C.text2 });
  });
  tbox(s, [{ text: "Despachar até", options: { fontSize: 16, breakLine: true } }, { text: "22h00", options: { fontSize: 40, bold: true } }],
    { x: 9.35, y: 5.65, w: 3.4, h: 1.05, fill: C.accent1, opts: { color: C.background1 } });
  s.addNotes("Exemplo com dados ilustrativos de um tanque de gasóleo de 30 mil litros. Durante a noite o consumo é baixo; de manhã desce mais depressa. Às 10h05 a plataforma deteta a entrega sozinha e confere: mediu 19 700 litros contra 20 000 faturados. Às 19h há 12 100 litros. A previsão mostra que chega ao nível crítico às 01h10 de quarta; como o camião precisa de 3 horas, a recomendação é despachar até às 22h00.");

  // =====================================================================
  // 13 · TRÊS CÁLCULOS
  // =====================================================================
  s = content("06 · CAMADA 4 — CÁLCULOS AUTOMÁTICOS", "Três números que a plataforma calcula sozinha", SEC);
  const calc = [
    ["MdSchedule", HEX.violet, "22h00", "Hora-limite para despachar", "Crítico previsto às 01h10 · o camião precisa de 3 h"],
    ["MdLocalShipping", HEX.amberDark, "−300 L", "Entrega abaixo da fatura (1,5%)", "Fatura 20 000 L · medido no tanque 19 700 L a 15 °C"],
    ["MdSearch", HEX.red, "−240 L", "Diferença no dia (2,4%)", "O tanque desceu 9 850 L · as bombas venderam 9 610 L"],
  ];
  calc.forEach(([n, col, big, t1, t2], i) => {
    const x = 0.6 + i * 4.15;
    box(s, { x, y: 1.75, w: 3.85, h: 4.4, fill: C.background2, radius: 0.14 });
    circleIcon(s, x + 0.35, 2.05, 0.8, n, col);
    T(s, big, { x: x + 0.35, y: 3.05, w: 3.3, h: 1.1, fontSize: 60, bold: true, color: col });
    T(s, t1, { x: x + 0.35, y: 4.2, w: 3.2, h: 0.75, fontSize: 18, bold: true });
    T(s, t2, { x: x + 0.35, y: 4.95, w: 3.2, h: 0.9, fontSize: 14, color: C.text2 });
  });
  T(s, "O terceiro cálculo (reconciliação diária) exige ler também as vendas das bombas — fase 2. As tolerâncias são definidas pela Sonangol.", { x: 0.6, y: 6.35, w: 12.1, h: 0.35, fontSize: 12, color: C.text2 });
  s.addNotes("Três cálculos automáticos. Primeiro, a hora-limite de despacho, a partir da previsão e do tempo de viagem. Segundo, a confirmação de entregas: a diferença entre fatura e volume medido, corrigido para 15 °C. Terceiro, a reconciliação diária: o que saiu do tanque contra o que as bombas venderam — diferenças acima do limite apontam para fuga, calibração ou desvio. Este último exige ler as vendas das bombas, numa fase 2.");

  // =====================================================================
  // 14 · ECRÃS
  // =====================================================================
  s = content("07 · MONITORIZAÇÃO", "O mesmo dado, no computador, no telemóvel e por SMS", SEC);
  // desktop
  box(s, { x: 0.6, y: 1.6, w: 7.0, h: 4.4, fill: C.background1, line: HEX.navy, lw: 2, radius: 0.1 });
  box(s, { x: 0.6, y: 1.6, w: 7.0, h: 0.42, fill: C.text1, round: false });
  T(s, "VIGIA · Despacho · Província de Luanda", { x: 0.8, y: 1.6, w: 5.5, h: 0.42, fontSize: 12, bold: true, fontFace: MONO, color: C.background1, valign: "middle" });
  const kpis = [["3", "postos críticos", HEX.red], ["11", "esgotam em 24 h", HEX.amberDark], ["7", "camiões em rota", HEX.violet], ["98%", "postos a reportar", HEX.green]];
  kpis.forEach(([a, b, col], i) => {
    const x = 0.8 + i * 1.68;
    box(s, { x, y: 2.2, w: 1.55, h: 0.95, fill: col, ft: 88, radius: 0.08 });
    T(s, a, { x: x + 0.12, y: 2.25, w: 1.35, h: 0.5, fontSize: 24, bold: true, color: col });
    T(s, b, { x: x + 0.12, y: 2.75, w: 1.4, h: 0.3, fontSize: 11, color: C.text1 });
  });
  const rows = [["Posto", "Esgota", "Ação"], ["Viana · T2 gasóleo", "~4 h", "DESPACHAR"], ["Cazenga · T1 gasolina", "~5 h", "DESPACHAR"], ["Talatona · T3 gasóleo", "~6 h", "DESPACHAR"], ["Rangel · T1 gasolina", "~10 h", "AGENDAR"], ["Cacuaco · T2 gasóleo", "~14 h", "MONITORIZAR"]];
  const actCol = { DESPACHAR: HEX.red, AGENDAR: HEX.amberDark, MONITORIZAR: HEX.green };
  s.addTable(rows.map((r, ri) => r.map((cell, ci) => ({
    text: cell,
    options: ri === 0 ? { bold: true, color: HEX.grey, fill: { color: HEX.panel }, fontSize: 11 }
      : { color: ci === 2 ? actCol[cell] : HEX.navy, bold: ci === 2, fontSize: 12, fontFace: ci === 1 ? MONO : undefined },
  }))), { x: 0.8, y: 3.35, w: 6.6, colW: [3.4, 1.4, 1.8], rowH: 0.36, border: { type: "solid", pt: 0.5, color: HEX.hair }, margin: [0.03, 0.08, 0.03, 0.08] });
  box(s, { x: 3.7, y: 6.0, w: 0.8, h: 0.25, fill: C.text1, round: false });
  box(s, { x: 3.1, y: 6.25, w: 2.0, h: 0.08, fill: C.text1, radius: 0.04 });
  T(s, "Computador — sede e despacho", { x: 0.6, y: 6.4, w: 7.0, h: 0.35, fontSize: 14, bold: true, align: "center" });
  // phone
  box(s, { x: 8.0, y: 1.6, w: 2.5, h: 4.75, fill: C.text1, radius: 0.3 });
  box(s, { x: 8.12, y: 1.85, w: 2.26, h: 4.3, fill: C.background2, radius: 0.2 });
  box(s, { x: 8.25, y: 2.0, w: 2.0, h: 1.0, fill: C.background1, line: HEX.hair, radius: 0.12 });
  T(s, "VIGIA · agora", { x: 8.38, y: 2.06, w: 1.8, h: 0.25, fontSize: 9, color: C.text2 });
  T(s, "CRÍTICO · Posto Viana", { x: 8.38, y: 2.3, w: 1.85, h: 0.3, fontSize: 11, bold: true, color: C.accent3 });
  T(s, "T2 gasóleo 9% · ~4 h", { x: 8.38, y: 2.6, w: 1.85, h: 0.3, fontSize: 11 });
  T(s, "Posto Viana", { x: 8.3, y: 3.15, w: 2.0, h: 0.35, fontSize: 14, bold: true });
  [[8.4, 0.67, HEX.green, "T1 67%"], [9.05, 0.09, HEX.red, "T2 9%"], [9.7, 0.42, HEX.amber, "T3 42%"]].forEach(([x, pct, col, lab]) => {
    box(s, { x, y: 3.6, w: 0.42, h: 1.4, fill: "DDE1E7", radius: 0.12 });
    box(s, { x, y: 3.6 + 1.4 * (1 - pct), w: 0.42, h: Math.max(1.4 * pct, 0.08), fill: col, radius: 0.08 });
    T(s, lab, { x: x - 0.15, y: 5.05, w: 0.72, h: 0.25, fontSize: 10, bold: col === HEX.red, color: col === HEX.red ? C.accent3 : C.text1, align: "center" });
  });
  tbox(s, "Confirmar despacho", { x: 8.25, y: 5.45, w: 2.0, h: 0.45, fill: C.accent1, opts: { fontSize: 12, bold: true, color: C.background1 } });
  T(s, "Telemóvel — gestor regional", { x: 7.7, y: 6.4, w: 3.1, h: 0.35, fontSize: 14, bold: true, align: "center" });
  // SMS
  circleIcon(s, 11.0, 1.65, 0.65, "MdSms", HEX.violet);
  T(s, "SMS · sem internet", { x: 11.0, y: 2.4, w: 1.8, h: 0.35, fontSize: 14, bold: true });
  tbox(s, "VIGIA CRITICO Viana T2 gasoleo 9% (~4h). Despachar ate 22:00. Resp 1=OK", { x: 11.0, y: 2.85, w: 1.75, h: 2.15, fill: C.background2, radius: 0.15, opts: { fontFace: MONO, fontSize: 11, align: "left", valign: "top", margin: 9 } });
  T(s, "Chega mesmo a quem não tem dados móveis", { x: 11.0, y: 5.15, w: 1.75, h: 0.75, fontSize: 12, color: C.text2 });
  s.addNotes("Os mesmos dados chegam a pessoas diferentes. A sede e o despacho veem a rede no computador: indicadores, mapa e a fila de reabastecimento por prioridade. O gestor regional recebe um alerta no telemóvel e confirma o despacho com um toque. E o SMS garante que o alerta crítico chega mesmo sem dados móveis. O acesso é por perfil: um gerente de posto só vê o seu posto. Ecrãs ilustrativos.");

  // =====================================================================
  // 15 · SEGURANÇA
  // =====================================================================
  SEC = "Segurança";
  pres.addSection({ title: SEC });
  s = content("08 · SEGURANÇA E FIABILIDADE", "Seguro por desenho, preparado para as condições angolanas", SEC);
  box(s, { x: 0.6, y: 1.65, w: 4.2, h: 5.05, fill: C.text1, radius: 0.14 });
  s.addShape(S.ellipse, { x: 0.95, y: 2.0, w: 1.0, h: 1.0, fill: { color: C.accent2 }, line: { type: "none" } });
  s.addImage({ data: ic("MdVisibility", "navy"), x: 1.18, y: 2.23, w: 0.54, h: 0.54 });
  T(s, "O sistema só lê.", { x: 0.95, y: 3.25, w: 3.6, h: 0.65, fontSize: 30, bold: true, color: C.background1 });
  T(s, "Sem acesso às bombas, aos preços, ao sistema de caixa ou a pagamentos. Numa rede separada da loja.", { x: 0.95, y: 4.0, w: 3.55, h: 1.5, fontSize: 16, color: "C9CFDD" });
  T(s, "Os dados são da Sonangol — alojados onde a Sonangol decidir.", { x: 0.95, y: 5.55, w: 3.55, h: 0.9, fontSize: 14, bold: true, color: C.accent2 });
  const sec = [
    ["MdLock", "Cifrado", "TLS ponta a ponta · APN privada opcional"],
    ["MdKey", "Identidade", "Certificado único por gateway; os falsos são rejeitados"],
    ["MdPeople", "Perfis", "Posto · região · sede, com autenticação de dois fatores"],
    ["MdBolt", "Cortes de energia", "Bateria 8–24 h e reinício automático"],
    ["MdWifiOff", "Falha de rede", "Troca de operador; guarda 30 dias de leituras"],
    ["MdWarning", "ATEX / IECEx", "Com ATG existente, nada novo dentro do tanque"],
  ];
  sec.forEach(([n, t1, t2], i) => {
    const x = 5.2 + (i % 3) * 2.53, y = 1.65 + Math.floor(i / 3) * 2.6;
    box(s, { x, y, w: 2.35, h: 2.45, fill: C.background2, radius: 0.12 });
    circleIcon(s, x + 0.25, y + 0.25, 0.65, n, i === 5 ? HEX.red : HEX.violet);
    T(s, t1, { x: x + 0.25, y: y + 1.0, w: 1.95, h: 0.4, fontSize: 16, bold: true });
    T(s, t2, { x: x + 0.25, y: y + 1.4, w: 1.95, h: 0.95, fontSize: 13, color: C.text2 });
  });
  s.addNotes("O ponto mais importante para a TI e a segurança: o sistema só lê. Não pode mexer nas bombas nem nos preços, porque não está ligado a eles. Os dados viajam cifrados, cada gateway tem identidade própria, e o acesso é por perfil com dois fatores. Para as condições angolanas: bateria para os cortes de energia, dois operadores e memória de 30 dias para falhas de rede. Na segurança física, nos postos com medidor existente não se instala nada dentro do tanque; qualquer sonda nova é certificada para atmosferas explosivas e instalada por técnico credenciado.");

  // =====================================================================
  // 16 · PILOTO
  // =====================================================================
  SEC = "Piloto e próximo passo";
  pres.addSection({ title: SEC });
  s = content("09 · PILOTO", "Piloto de 90 dias · 10–20 postos em Luanda", SEC);
  const gx = 4.0, gw = 8.75, wk = gw / 12;
  for (let w = 0; w < 12; w++) {
    if (w % 2 === 0) box(s, { x: gx + w * wk, y: 1.95, w: wk, h: 3.85, fill: C.background2, round: false });
    T(s, "S" + (w + 1), { x: gx + w * wk, y: 1.62, w: wk, h: 0.28, fontSize: 11, bold: true, color: C.text2, align: "center" });
  }
  const gantt = [
    ["Levantamento dos postos", 1, 2, C.text1],
    ["Compra e preparação", 2, 4, C.text2],
    ["Instalação nos postos", 3, 6, C.accent1],
    ["Painel ao vivo + alertas", 6, 12, C.accent5],
    ["Modelo de previsão", 9, 12, C.accent2],
    ["Avaliação e relatório", 11, 12, C.accent4],
  ];
  gantt.forEach(([n, a, b, col], i) => {
    const y = 2.05 + i * 0.6;
    T(s, n, { x: 0.6, y, w: 3.3, h: 0.42, fontSize: 15, bold: i === 0, valign: "middle" });
    box(s, { x: gx + (a - 1) * wk + 0.03, y, w: (b - a + 1) * wk - 0.06, h: 0.42, fill: col, radius: 0.08 });
  });
  [[6, "1.ª leitura real no painel"], [12, "Decisão de expansão"]].forEach(([w, lab], i) => {
    const x = gx + w * wk;
    s.addShape(S.diamond, { x: x - 0.15, y: 5.72, w: 0.3, h: 0.3, fill: { color: C.accent3 }, line: { type: "none" } });
    T(s, lab, { x: i === 0 ? x - 1.4 : x - 2.8, y: 6.05, w: 2.8, h: 0.3, fontSize: 12, bold: true, color: C.accent3, align: i === 0 ? "center" : "right" });
  });
  T(s, "MEDIMOS", { x: 0.6, y: 6.4, w: 1.2, h: 0.35, fontSize: 11, bold: true, fontFace: MONO, color: C.accent1, valign: "middle" });
  ["Horas de posto seco", "Ruturas e emergências", "Erro da previsão", "Fatura vs. recebido"].forEach((m, i) =>
    tbox(s, m, { x: 1.75 + i * 2.75, y: 6.4, w: 2.55, h: 0.38, fill: C.accent1, ft: 86, radius: 0.19, opts: { fontSize: 13, bold: true } }));
  s.addNotes("Proposta concreta: piloto de 90 dias em 10 a 20 postos de Luanda — por exemplo Talatona, Kilamba Kiaxi e Viana — com postos com e sem medidor automático, para testar os dois cenários. Semanas 1 e 2: levantamento. Até à semana 6: instalação e primeira leitura real no painel. Depois, painel ao vivo, modelo de previsão e avaliação. Medimos horas de posto seco, ruturas e entregas de emergência, o erro da previsão e as diferenças entre faturas e volumes recebidos — antes e durante o piloto.");

  // =====================================================================
  // 17 · CUSTOS
  // =====================================================================
  s = content("09 · CUSTOS ESTIMADOS", "Custos por posto, em ordem de grandeza", SEC);
  const cost = [
    ["MdRouter", "US$ 150–500", "Gateway 4G + antena", "Uma vez, por posto"],
    ["MdBatteryChargingFull", "US$ 50–150", "UPS / bateria dedicada", "Só se o posto não tiver"],
    ["MdSignalCellularAlt", "Poucos US$/mês", "2 cartões SIM M2M", "≈ 5 MB de dados por mês"],
  ];
  cost.forEach(([n, big, t1, t2], i) => {
    const x = 0.6 + i * 4.15;
    box(s, { x, y: 1.7, w: 3.85, h: 2.55, fill: C.background2, radius: 0.14 });
    circleIcon(s, x + 0.3, 1.95, 0.7, n, HEX.violet);
    T(s, big, { x: x + 0.3, y: 2.8, w: 3.4, h: 0.65, fontSize: 30, bold: true, color: C.accent1 });
    T(s, t1, { x: x + 0.3, y: 3.42, w: 3.4, h: 0.38, fontSize: 16, bold: true });
    T(s, t2, { x: x + 0.3, y: 3.8, w: 3.4, h: 0.35, fontSize: 14, color: C.text2 });
  });
  box(s, { x: 0.6, y: 4.5, w: 12.15, h: 1.05, fill: C.accent2, ft: 82, radius: 0.12 });
  circleIcon(s, 0.85, 4.68, 0.7, "MdOilBarrel", HEX.amberDark);
  T(s, [{ text: "Postos sem medidor automático: ", options: { bold: true } }, { text: "sonda ou radar certificado + instalação — de centenas a mais de US$ 2 000 por tanque." }],
    { x: 1.75, y: 4.5, w: 10.8, h: 1.05, fontSize: 16, valign: "middle" });
  box(s, { x: 0.6, y: 5.8, w: 12.15, h: 0.9, fill: C.accent3, ft: 88, radius: 0.12 });
  s.addImage({ data: ic("MdWarning", "red"), x: 0.85, y: 5.98, w: 0.5, h: 0.5 });
  T(s, [{ text: "Estimativas de mercado, não uma cotação. ", options: { bold: true, color: C.accent3 } }, { text: "Os valores reais saem do levantamento: fornecedores locais, impostos e importação." }],
    { x: 1.55, y: 5.8, w: 11.0, h: 0.9, fontSize: 15, valign: "middle" });
  s.addNotes("Ordem de grandeza, para orientação: o gateway com antena custa entre 150 e 500 dólares por posto, uma vez. Uma UPS dedicada, se o posto não tiver, 50 a 150. Os dados móveis são poucos dólares por mês. O custo grande está nos postos sem medidor automático, onde é preciso instalar uma sonda certificada. Sublinho: são estimativas, não uma cotação — os números reais saem do levantamento.");

  // =====================================================================
  // 18 · PRÓXIMO PASSO
  // =====================================================================
  s = pres.addSlide({ masterName: "VIGIA_DARK", sectionTitle: SEC });
  s.addText("10 · PRÓXIMO PASSO", { placeholder: "eyebrow" });
  s.addText("Visita técnica a 2–3 postos em Luanda", { placeholder: "title" });
  T(s, "Fotografar as consolas, medir o sinal móvel e confirmar a ligação — e o plano passa de estimativa a proposta concreta.", { x: 0.6, y: 1.95, w: 11.5, h: 0.8, fontSize: 18, color: "C9CFDD" });
  const qs = ["Que marcas de medidor automático (ATG) estão instaladas nos postos?", "Onde alojar a plataforma: datacenter da Sonangol ou nuvem?", "Quem lidera o piloto: Operações, TI ou Logística?"];
  qs.forEach((qq, i) => {
    const x = 0.6 + i * 4.15;
    box(s, { x, y: 3.1, w: 3.85, h: 2.4, fill: C.background1, ft: 92, line: "3A4A6B", radius: 0.14 });
    tbox(s, String(i + 1), { x: x + 0.3, y: 3.4, w: 0.7, h: 0.7, fill: C.accent2, opts: { shape: S.ellipse, fontSize: 22, bold: true, color: C.text1, margin: 0 } });
    T(s, qq, { x: x + 0.3, y: 4.3, w: 3.3, h: 1.1, fontSize: 18, bold: true, color: C.background1 });
  });
  T(s, [{ text: "Aires Augusto Miguêns", options: { bold: true } }, { text: "  ·  airesmiguens@gmail.com  ·  +244 931 391 063  ·  airestech.us" }],
    { x: 0.6, y: 6.3, w: 9.6, h: 0.45, fontSize: 16, color: C.background1 });
  s.addImage({ data: LOGO_W, x: 12.75 - 0.5 * LOGO_RATIO, y: 6.28, w: 0.5 * LOGO_RATIO, h: 0.5, altText: "Sonangol", objectName: "logo-sonangol" });
  s.addNotes("Proponho como próximo passo uma visita técnica a dois ou três postos em Luanda, com um técnico da Sonangol. Com isso respondemos às três perguntas que decidem o desenho do piloto: que medidores existem, onde alojar a plataforma, e quem lidera do lado da Sonangol.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
