const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "Last Minute";
pptx.company = "Last Minute";
pptx.subject = "Bilingual company profile";
pptx.title = "Last Minute Company Profile";
pptx.lang = "ar-SA";
pptx.theme = {
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  lang: "ar-SA"
};
pptx.defineLayout({ name: "CUSTOM_WIDE", width: 13.333, height: 7.5 });
pptx.layout = "CUSTOM_WIDE";

const root = path.resolve(__dirname, "..");
const assets = path.join(root, "dist", "assets");
const out = path.join(root, "dist", "downloads", "Last-Minute-Company-Profile.pptx");

const LOGO_LIGHT = path.join(assets, "logo-light.svg");
const LOGO_DARK = path.join(assets, "logo-dark.svg");
const PATTERN = path.join(assets, "brand-pattern.svg");
const HERO = path.join(assets, "event-operations-hero.png");

const C = {
  navy: "0A192F",
  navyDark: "061426",
  navyMid: "112A45",
  teal: "09B997",
  tealDark: "07846F",
  tealPale: "D9F7F0",
  paper: "F4F7F8",
  white: "FFFFFF",
  ink: "0F2133",
  muted: "5F6F7E",
  line: "D8E0E5",
  soft: "E9EEF1"
};

const SH = pptx.shapes;

function addText(slide, text, opts = {}) {
  slide.addText(text, {
    fontFace: "Arial",
    margin: 0,
    breakLine: false,
    fit: "shrink",
    valign: "mid",
    ...opts
  });
}

function addAr(slide, text, opts = {}) {
  addText(slide, text, {
    rtlMode: true,
    align: "right",
    fontFace: "Arial",
    ...opts
  });
}

function addEn(slide, text, opts = {}) {
  addText(slide, text, {
    rtlMode: false,
    align: "left",
    fontFace: "Arial",
    ...opts
  });
}

function addBrandPattern(slide, options = {}) {
  slide.addImage({
    path: PATTERN,
    x: options.x ?? 0,
    y: options.y ?? 0,
    w: options.w ?? 5.2,
    h: options.h ?? 7.5,
    transparency: options.transparency ?? 86,
    altText: "Last Minute brand pattern"
  });
}

function addLogo(slide, light = false, x = 0.55, y = 0.34, w = 2.15) {
  slide.addImage({
    path: light ? LOGO_LIGHT : LOGO_DARK,
    x,
    y,
    w,
    h: w * 0.371,
    altText: "Last Minute logo"
  });
}

function addSlideChrome(slide, number, dark = false) {
  addLogo(slide, dark, 0.55, 0.32, 1.6);
  addEn(slide, String(number).padStart(2, "0"), {
    x: 12.2,
    y: 0.42,
    w: 0.52,
    h: 0.22,
    fontSize: 9,
    bold: true,
    color: dark ? C.teal : C.tealDark,
    align: "right",
    charSpacing: 2
  });
}

function addSectionTitle(slide, ar, en, options = {}) {
  const dark = options.dark ?? false;
  const arX = options.arX ?? 6.92;
  const enX = options.enX ?? 0.72;
  const y = options.y ?? 1.25;
  addAr(slide, ar, {
    x: arX,
    y,
    w: 5.7,
    h: 0.78,
    fontSize: options.arSize ?? 27,
    bold: true,
    color: dark ? C.white : C.ink
  });
  addEn(slide, en.toUpperCase(), {
    x: enX,
    y: y + 0.1,
    w: 5.6,
    h: 0.48,
    fontSize: options.enSize ?? 14,
    bold: true,
    color: dark ? C.teal : C.tealDark,
    charSpacing: 1.2
  });
}

function addDivider(slide, dark = false, y = 2.18) {
  slide.addShape(SH.LINE, {
    x: 0.72,
    y,
    w: 11.9,
    h: 0,
    line: { color: dark ? "36506A" : C.line, width: 1 }
  });
}

function addBilingualParagraph(slide, ar, en, options = {}) {
  const y = options.y ?? 2.55;
  const h = options.h ?? 2.1;
  const arX = options.arX ?? 7.05;
  const enX = options.enX ?? 0.72;
  addAr(slide, ar, {
    x: arX,
    y,
    w: 5.55,
    h,
    fontSize: options.arSize ?? 20,
    color: options.arColor ?? C.ink,
    breakLine: true,
    valign: "top",
    lineSpacingMultiple: 1.08
  });
  addEn(slide, en, {
    x: enX,
    y,
    w: 5.5,
    h,
    fontSize: options.enSize ?? 15,
    color: options.enColor ?? C.muted,
    breakLine: true,
    valign: "top",
    lineSpacingMultiple: 1.08
  });
}

function addBottomTag(slide, ar, en, dark = false) {
  const fill = dark ? C.teal : C.navy;
  const color = dark ? C.navyDark : C.white;
  slide.addShape(SH.RECTANGLE, {
    x: 0.72,
    y: 6.58,
    w: 11.9,
    h: 0.42,
    fill: { color: fill },
    line: { color: fill }
  });
  addAr(slide, ar, {
    x: 6.72,
    y: 6.67,
    w: 5.55,
    h: 0.18,
    fontSize: 10.3,
    bold: true,
    color
  });
  addEn(slide, en.toUpperCase(), {
    x: 1.05,
    y: 6.67,
    w: 5.5,
    h: 0.18,
    fontSize: 8.8,
    bold: true,
    color,
    charSpacing: 1
  });
}

function addCard(slide, { x, y, w, h, number, arTitle, enTitle, arText, enText, dark = false, teal = false }) {
  const fill = teal ? C.teal : dark ? C.navy : C.white;
  const titleColor = teal ? C.navyDark : dark ? C.white : C.ink;
  const bodyColor = teal ? "153C48" : dark ? "DCE4EA" : "435768";
  slide.addShape(SH.RECTANGLE, {
    x, y, w, h,
    fill: { color: fill },
    line: { color: dark ? "29415A" : teal ? C.teal : C.line, width: 1 }
  });
  addEn(slide, number, {
    x: x + 0.25,
    y: y + 0.22,
    w: 0.7,
    h: 0.45,
    fontSize: 22,
    bold: true,
    color: teal ? C.navyDark : C.teal
  });
  addAr(slide, arTitle, {
    x: x + w * 0.5,
    y: y + 0.27,
    w: w * 0.44,
    h: 0.55,
    fontSize: 18,
    bold: true,
    color: titleColor
  });
  addEn(slide, enTitle.toUpperCase(), {
    x: x + 0.25,
    y: y + 0.85,
    w: w - 0.5,
    h: 0.38,
    fontSize: 10,
    bold: true,
    color: teal ? C.navyDark : C.teal,
    charSpacing: 0.7
  });
  addAr(slide, arText, {
    x: x + 0.28,
    y: y + 1.32,
    w: w - 0.56,
    h: h - 1.62,
    fontSize: 13.7,
    color: bodyColor,
    valign: "top"
  });
  addEn(slide, enText, {
    x: x + 0.28,
    y: y + h - 0.72,
    w: w - 0.56,
    h: 0.42,
    fontSize: 9.2,
    italic: true,
    color: bodyColor,
    valign: "top"
  });
}

// 01 — Cover
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navyDark };
  slide.addImage({ path: HERO, x: 0, y: 0, w: 13.333, h: 7.5, altText: "Event operations team coordinating an urgent delivery" });
  slide.addShape(SH.RECTANGLE, {
    x: 0,
    y: 0,
    w: 8.2,
    h: 7.5,
    fill: { color: C.navyDark, transparency: 5 },
    line: { color: C.navyDark, transparency: 100 }
  });
  slide.addShape(SH.RECTANGLE, {
    x: 7.9,
    y: 0,
    w: 1.6,
    h: 7.5,
    fill: { color: C.navyDark, transparency: 45 },
    line: { color: C.navyDark, transparency: 100 }
  });
  addBrandPattern(slide, { x: 0, y: 0, w: 4.9, h: 7.5, transparency: 89 });
  addLogo(slide, true, 0.75, 0.55, 2.7);
  addAr(slide, "الملف التعريفي", {
    x: 0.78,
    y: 2.2,
    w: 6.5,
    h: 0.7,
    fontSize: 31,
    color: C.teal,
    bold: true
  });
  addAr(slide, "وساطة توريد مرنة\nللفعاليات والشركات", {
    x: 0.78,
    y: 2.95,
    w: 6.5,
    h: 1.6,
    fontSize: 31,
    color: C.white,
    bold: true,
    breakLine: true,
    valign: "top"
  });
  addEn(slide, "FLEXIBLE SOURCING & SUPPLIER BROKERAGE", {
    x: 0.8,
    y: 4.85,
    w: 6.4,
    h: 0.35,
    fontSize: 11,
    bold: true,
    color: C.teal,
    charSpacing: 1.3
  });
  addEn(slide, "Company Profile", {
    x: 0.8,
    y: 5.27,
    w: 4.5,
    h: 0.48,
    fontSize: 18,
    color: C.white
  });
  addEn(slide, "2026", {
    x: 0.8,
    y: 6.65,
    w: 0.7,
    h: 0.22,
    fontSize: 9,
    bold: true,
    color: C.white,
    charSpacing: 1
  });
}

// 02 — Executive summary
{
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addSlideChrome(slide, 2, false);
  addSectionTitle(slide, "من نحن", "Who we are");
  addDivider(slide);
  addBilingualParagraph(
    slide,
    "نحن شركة وساطة وتوريد مرن نربط منظمي الفعاليات والشركات بشبكة من الموردين. نتدخل عندما يصبح الوقت ضيقًا، أو عندما يتوفر المنتج خارج مسار الشراء المعتاد، لننسق البحث والعرض والشراء والفوترة والمتابعة حتى التسليم.",
    "Last Minute is a flexible sourcing and brokerage company connecting event organizers and businesses with a supplier network. We step in when time is limited or when the right product sits outside the usual procurement route, coordinating sourcing, quotation, purchasing, invoicing and follow-through to delivery.",
    { y: 2.55, h: 2.55, arSize: 18.5, enSize: 15 }
  );
  slide.addShape(SH.RECTANGLE, { x: 0.72, y: 5.52, w: 3.62, h: 0.8, fill: { color: C.white }, line: { color: C.line } });
  slide.addShape(SH.RECTANGLE, { x: 4.85, y: 5.52, w: 3.62, h: 0.8, fill: { color: C.white }, line: { color: C.line } });
  slide.addShape(SH.RECTANGLE, { x: 8.98, y: 5.52, w: 3.64, h: 0.8, fill: { color: C.navy }, line: { color: C.navy } });
  addAr(slide, "سرعة الاستجابة", { x: 2.0, y: 5.69, w: 2.05, h: 0.24, fontSize: 13, bold: true, color: C.ink });
  addEn(slide, "RESPONSIVE", { x: 0.98, y: 5.72, w: 1.0, h: 0.18, fontSize: 8, bold: true, color: C.tealDark });
  addAr(slide, "تنسيق واضح", { x: 6.12, y: 5.69, w: 2.05, h: 0.24, fontSize: 13, bold: true, color: C.ink });
  addEn(slide, "CLEAR", { x: 5.12, y: 5.72, w: 0.8, h: 0.18, fontSize: 8, bold: true, color: C.tealDark });
  addAr(slide, "شبكة مرنة", { x: 10.23, y: 5.69, w: 2.05, h: 0.24, fontSize: 13, bold: true, color: C.white });
  addEn(slide, "FLEXIBLE", { x: 9.25, y: 5.72, w: 0.85, h: 0.18, fontSize: 8, bold: true, color: C.teal });
}

// 03 — Purpose, mission, vision
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navyDark };
  addBrandPattern(slide, { x: 8.8, y: 0, w: 4.53, h: 7.5, transparency: 96 });
  addSlideChrome(slide, 3, true);
  addSectionTitle(slide, "وجهتنا", "Purpose · Mission · Vision", { dark: true });
  addDivider(slide, true);
  const items = [
    ["01", "غايتنا", "PURPOSE", "إزالة العوائق التي تمنع وصول المنتج المناسب إلى موقع التنفيذ في الوقت المناسب.", "Remove the friction that keeps the right product from reaching the right place on time."],
    ["02", "مهمتنا", "MISSION", "تبسيط التوريد العاجل والشراء المؤسسي عبر شبكة موردين وتنسيق تجاري واضح.", "Simplify urgent sourcing and corporate purchasing through a supplier network and clear coordination."],
    ["03", "رؤيتنا", "VISION", "أن نكون حلقة الوصل الموثوقة للطلبات العاجلة في قطاع الفعاليات والأعمال.", "Become the trusted connection for urgent sourcing across events and business operations."]
  ];
  items.forEach((item, i) => {
    const y = 2.55 + i * 1.26;
    addEn(slide, item[0], { x: 0.75, y, w: 0.5, h: 0.35, fontSize: 17, bold: true, color: C.teal });
    addEn(slide, item[2], { x: 1.55, y: y + 0.02, w: 1.1, h: 0.22, fontSize: 9, bold: true, color: C.teal, charSpacing: 0.8 });
    addEn(slide, item[4], { x: 2.7, y: y - 0.03, w: 3.38, h: 0.56, fontSize: 12.1, color: "CDD8E0", valign: "top" });
    addAr(slide, item[1], { x: 10.45, y: y - 0.03, w: 1.8, h: 0.35, fontSize: 18, bold: true, color: C.white });
    addAr(slide, item[3], { x: 6.65, y: y - 0.03, w: 3.55, h: 0.62, fontSize: 14.3, color: "E8EEF2", valign: "top" });
    if (i < 2) slide.addShape(SH.LINE, { x: 0.75, y: y + 0.8, w: 11.5, h: 0, line: { color: "29415A", width: 1 } });
  });
  addBottomTag(slide, "وقت أقل في البحث. وضوح أكبر في التنفيذ.", "Less search time. More execution clarity.", true);
}

// 04 — Two problems
{
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addSlideChrome(slide, 4, false);
  addSectionTitle(slide, "المشكلتان اللتان نعالجهما", "The two problems we solve", { arSize: 25 });
  addDivider(slide);
  addCard(slide, {
    x: 0.72, y: 2.48, w: 5.75, h: 3.7, number: "01",
    arTitle: "تغيير في اللحظة الأخيرة",
    enTitle: "Last-minute change",
    arText: "زيادة كمية، استبدال منتج، أو تعديل مواصفات في وقت يصعب فيه على المنظم إيجاد مورد متاح. يتحول التأخير هنا إلى خطر مباشر على جودة الفعالية.",
    enText: "A quantity increase, replacement or specification change can become an immediate delivery risk.",
    dark: true
  });
  addCard(slide, {
    x: 6.86, y: 2.48, w: 5.75, h: 3.7, number: "02",
    arTitle: "فجوة الشراء والفوترة",
    enTitle: "Purchasing & invoicing gap",
    arText: "قد يكون المنتج متاحًا لدى متجر أو مورد لا يتوافق مباشرة مع إجراءات مشتريات الشركة. وجود المنتج وحده لا يكفي إذا تعذر شراؤه بطريقة مناسبة.",
    enText: "The product may exist, yet remain inaccessible through the organizer’s normal procurement process.",
    teal: true
  });
}

// 05 — Urgent change detail
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideChrome(slide, 5, false);
  addSectionTitle(slide, "عندما يتغيّر الطلب", "When the requirement changes");
  addDivider(slide);
  const cases = [
    ["زيادة", "QUANTITY", "زيادة عدد الضيوف أو الوحدات المطلوبة.", "More guests or units than originally planned."],
    ["استبدال", "REPLACEMENT", "عدم توفر المنتج الأساسي والحاجة إلى بديل مناسب.", "The original item is unavailable and a suitable alternative is needed."],
    ["تخصيص", "SPECIFICATION", "تعديل المقاس أو اللون أو التفاصيل التشغيلية.", "A size, color or operating detail changes."],
    ["وقت", "DEADLINE", "تقليص المهلة المتاحة قبل التنفيذ.", "The available lead time suddenly becomes shorter."]
  ];
  cases.forEach((c, i) => {
    const x = 0.72 + i * 3.0;
    slide.addShape(SH.RECTANGLE, { x, y: 2.55, w: 2.68, h: 3.55, fill: { color: i === 3 ? C.navy : C.paper }, line: { color: i === 3 ? C.navy : C.line } });
    addEn(slide, String(i + 1).padStart(2, "0"), { x: x + 0.25, y: 2.82, w: 0.5, h: 0.3, fontSize: 15, bold: true, color: C.teal });
    addAr(slide, c[0], { x: x + 0.25, y: 3.35, w: 2.15, h: 0.45, fontSize: 21, bold: true, color: i === 3 ? C.white : C.ink });
    addEn(slide, c[1], { x: x + 0.25, y: 3.9, w: 2.15, h: 0.22, fontSize: 8.5, bold: true, color: C.teal, charSpacing: 0.7 });
    addAr(slide, c[2], { x: x + 0.25, y: 4.4, w: 2.15, h: 0.76, fontSize: 12.8, color: i === 3 ? "E1E9EE" : "435768", valign: "top" });
    addEn(slide, c[3], { x: x + 0.25, y: 5.32, w: 2.15, h: 0.5, fontSize: 8.8, color: i === 3 ? "C3D0D9" : "435768", italic: true, valign: "top" });
  });
  addBottomTag(slide, "نبحث عن الخيار المناسب ضمن الوقت المتاح، لا ضمن الظروف المثالية.", "We source for the time available, not for ideal conditions.");
}

// 06 — Procurement gap detail
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navyDark };
  addBrandPattern(slide, { x: 0, y: 0.8, w: 5.2, h: 6.7, transparency: 96 });
  addSlideChrome(slide, 6, true);
  addSectionTitle(slide, "عندما يكون المنتج موجودًا لكن الشراء معقّد", "When the item exists but purchasing is difficult", { dark: true, arSize: 24, enSize: 10.8 });
  addDivider(slide, true);
  const nodes = [
    { x: 0.82, ar: "المورد", en: "SUPPLIER", textAr: "منتج متاح", textEn: "Item available" },
    { x: 4.86, ar: "Last Minute", en: "COORDINATION", textAr: "تنسيق تجاري مركزي", textEn: "Central commercial route" },
    { x: 8.9, ar: "المنظم", en: "ORGANIZER", textAr: "شراء وفاتورة واضحة", textEn: "Clear purchase & invoice" }
  ];
  nodes.forEach((n, i) => {
    slide.addShape(SH.RECTANGLE, { x: n.x, y: 2.82, w: 3.35, h: 2.35, fill: { color: i === 1 ? C.teal : C.navyMid }, line: { color: i === 1 ? C.teal : "36506A" } });
    addAr(slide, n.ar, { x: n.x + 0.28, y: 3.12, w: 2.78, h: 0.4, fontSize: 19, bold: true, color: i === 1 ? C.navyDark : C.white });
    addEn(slide, n.en, { x: n.x + 0.28, y: 3.7, w: 2.78, h: 0.22, fontSize: 8.5, bold: true, color: i === 1 ? C.navyDark : C.teal, charSpacing: 0.8 });
    addAr(slide, n.textAr, { x: n.x + 0.28, y: 4.24, w: 2.78, h: 0.3, fontSize: 12.5, color: i === 1 ? "153C48" : "C7D2DA" });
    addEn(slide, n.textEn, { x: n.x + 0.28, y: 4.63, w: 2.78, h: 0.25, fontSize: 9, color: i === 1 ? "153C48" : "C5D1DA" });
    if (i < 2) {
      slide.addShape(SH.LINE, { x: n.x + 3.35, y: 3.98, w: 0.69, h: 0, line: { color: C.teal, width: 3, beginArrowType: "none", endArrowType: "triangle" } });
    }
  });
  addBottomTag(slide, "نحوّل التوفر في السوق إلى مسار شراء قابل للتنفيذ.", "We turn market availability into an actionable purchasing route.", true);
}

// 07 — How it works
{
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addSlideChrome(slide, 7, false);
  addSectionTitle(slide, "كيف نعمل", "How it works");
  addDivider(slide);
  const steps = [
    ["01", "استلام الاحتياج", "RECEIVE", "المنتج، الكمية، المواصفات، الموقع، والوقت."],
    ["02", "البحث والمطابقة", "MATCH", "مراجعة الموردين والتوفر والبدائل الممكنة."],
    ["03", "العرض والتنسيق", "COORDINATE", "توضيح السعر والمدة ومسار الشراء والفوترة."],
    ["04", "التوريد والمتابعة", "DELIVER", "متابعة التنفيذ حتى نقطة التسليم."]
  ];
  slide.addShape(SH.LINE, { x: 1.3, y: 3.34, w: 10.55, h: 0, line: { color: C.line, width: 2 } });
  steps.forEach((s, i) => {
    const x = 0.72 + i * 3.02;
    slide.addShape(SH.OVAL, { x: x + 0.75, y: 2.82, w: 1.04, h: 1.04, fill: { color: C.white }, line: { color: C.teal, width: 2 } });
    addEn(slide, s[0], { x: x + 0.75, y: 3.12, w: 1.04, h: 0.22, fontSize: 12, bold: true, color: C.navy, align: "center" });
    addAr(slide, s[1], { x, y: 4.25, w: 2.52, h: 0.4, fontSize: 17, bold: true, color: C.ink, align: "center" });
    addEn(slide, s[2], { x, y: 4.78, w: 2.52, h: 0.2, fontSize: 8.2, bold: true, color: C.tealDark, align: "center", charSpacing: 0.6 });
    addAr(slide, s[3], { x, y: 5.18, w: 2.52, h: 0.68, fontSize: 12.4, color: "435768", align: "center", valign: "top" });
  });
  addBottomTag(slide, "نقطة اتصال واحدة من الطلب إلى التسليم.", "One point of coordination from request to delivery.");
}

// 08 — Services
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideChrome(slide, 8, false);
  addSectionTitle(slide, "خدماتنا", "Our services");
  addDivider(slide);
  const services = [
    ["A", "التوريد العاجل", "URGENT SOURCING", "البحث عن المنتجات والكميات المطلوبة ضمن الوقت المتاح."],
    ["B", "وساطة الموردين", "SUPPLIER BROKERAGE", "ربط الاحتياج بالمورد الأنسب من حيث التوفر والموقع."],
    ["C", "تنسيق الشراء والفوترة", "PURCHASE COORDINATION", "تسهيل المسار التجاري عندما لا تلائم الخيارات إجراءات الشركة مباشرة."],
    ["D", "متابعة التنفيذ", "EXECUTION FOLLOW-UP", "تحديثات واضحة ومتابعة حتى الوصول إلى نقطة التسليم."]
  ];
  services.forEach((s, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.72 + col * 6.14;
    const y = 2.48 + row * 1.78;
    slide.addShape(SH.RECTANGLE, { x, y, w: 5.75, h: 1.48, fill: { color: i === 3 ? C.navy : C.paper }, line: { color: i === 3 ? C.navy : C.line } });
    addEn(slide, s[0], { x: x + 0.3, y: y + 0.25, w: 0.42, h: 0.32, fontSize: 18, bold: true, color: C.teal });
    addAr(slide, s[1], { x: x + 2.55, y: y + 0.22, w: 2.85, h: 0.35, fontSize: 16.5, bold: true, color: i === 3 ? C.white : C.ink });
    addEn(slide, s[2], { x: x + 0.88, y: y + 0.3, w: 1.9, h: 0.2, fontSize: 7.8, bold: true, color: C.teal, charSpacing: 0.45 });
    addAr(slide, s[3], { x: x + 0.3, y: y + 0.78, w: 5.1, h: 0.48, fontSize: 12.6, color: i === 3 ? "D9E3E9" : "435768", valign: "top" });
  });
  addBottomTag(slide, "دعم عملي قبل الفعالية وأثناءها.", "Practical support before and during the event.");
}

// 09 — Value ecosystem
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navyDark };
  addBrandPattern(slide, { x: 8.9, y: 0.2, w: 4.4, h: 7.1, transparency: 91 });
  addSlideChrome(slide, 9, true);
  addSectionTitle(slide, "قيمة لكل طرف", "Value for every side", { dark: true });
  addDivider(slide, true);
  const columns = [
    ["المنظم", "ORGANIZER", ["استجابة للطلبات الطارئة", "خيارات أكثر", "وقت أقل في البحث"]],
    ["المورد", "SUPPLIER", ["قناة مبيعات جديدة", "عملاء من قطاع الأعمال", "فرص توريد متكررة"]],
    ["Last Minute", "COORDINATOR", ["عمولة أو هامش خدمة", "نموذج مرن", "رؤية أفضل للأداء"]]
  ];
  columns.forEach((c, i) => {
    const x = 0.72 + i * 4.05;
    slide.addShape(SH.RECTANGLE, { x, y: 2.55, w: 3.72, h: 3.55, fill: { color: i === 2 ? C.teal : C.navyMid }, line: { color: i === 2 ? C.teal : "36506A" } });
    addAr(slide, c[0], { x: x + 0.3, y: 2.9, w: 3.12, h: 0.45, fontSize: 21, bold: true, color: i === 2 ? C.navyDark : C.white });
    addEn(slide, c[1], { x: x + 0.3, y: 3.5, w: 3.12, h: 0.22, fontSize: 8.5, bold: true, color: i === 2 ? C.navyDark : C.teal, charSpacing: 0.75 });
    c[2].forEach((item, index) => {
      slide.addShape(SH.LINE, { x: x + 0.3, y: 4.12 + index * 0.58, w: 0.32, h: 0, line: { color: i === 2 ? C.navyDark : C.teal, width: 2 } });
      addAr(slide, item, { x: x + 0.75, y: 3.96 + index * 0.58, w: 2.65, h: 0.3, fontSize: 13.3, color: i === 2 ? "153C48" : "E5ECF0" });
    });
  });
  addBottomTag(slide, "نزيل الاحتكاك بين الاحتياج والتوفر.", "Removing friction between demand and availability.", true);
}

// 10 — Business model
{
  const slide = pptx.addSlide();
  slide.background = { color: C.paper };
  addSlideChrome(slide, 10, false);
  addSectionTitle(slide, "نموذج العمل", "Business model");
  addDivider(slide);
  const revenue = [
    ["هامش الخدمة العاجلة", "URGENT SERVICE MARGIN", "يعكس أولوية الطلب وسرعة البحث والتنسيق."],
    ["عمولة الوساطة", "BROKERAGE COMMISSION", "على عمليات التوريد التي تتم عبر شبكة الموردين."],
    ["رسوم التنسيق", "COORDINATION FEE", "عند إدارة مسار الشراء والفوترة والمتابعة."]
  ];
  revenue.forEach((r, i) => {
    const y = 2.55 + i * 1.22;
    slide.addShape(SH.RECTANGLE, { x: 0.72, y, w: 11.9, h: 0.92, fill: { color: i === 0 ? C.navy : C.white }, line: { color: i === 0 ? C.navy : C.line } });
    addEn(slide, String(i + 1).padStart(2, "0"), { x: 0.98, y: y + 0.29, w: 0.42, h: 0.2, fontSize: 10, bold: true, color: C.teal });
    addEn(slide, r[1], { x: 1.65, y: y + 0.29, w: 2.6, h: 0.2, fontSize: 8.5, bold: true, color: C.teal, charSpacing: 0.5 });
    addAr(slide, r[0], { x: 9.1, y: y + 0.21, w: 3.1, h: 0.35, fontSize: 17, bold: true, color: i === 0 ? C.white : C.ink });
    addAr(slide, r[2], { x: 4.42, y: y + 0.19, w: 4.2, h: 0.4, fontSize: 13, color: i === 0 ? "DCE5EB" : "435768" });
  });
  slide.addShape(SH.RECTANGLE, { x: 0.72, y: 6.28, w: 11.9, h: 0.42, fill: { color: C.tealPale }, line: { color: C.tealPale } });
  addAr(slide, "يعتمد السعر النهائي على طبيعة الطلب، درجة الاستعجال، التوفر، وموقع التسليم.", { x: 6.55, y: 6.37, w: 5.65, h: 0.2, fontSize: 10, color: C.tealDark });
  addEn(slide, "Final pricing depends on urgency, availability, request details and delivery location.", { x: 1.05, y: 6.37, w: 5.35, h: 0.2, fontSize: 9.2, color: C.tealDark });
}

// 11 — Sectors
{
  const slide = pptx.addSlide();
  slide.background = { color: C.navyDark };
  slide.addImage({ path: HERO, x: 0, y: 0, w: 13.333, h: 7.5, transparency: 18, altText: "Corporate event setup" });
  slide.addShape(SH.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: C.navyDark, transparency: 14 }, line: { color: C.navyDark, transparency: 100 } });
  addSlideChrome(slide, 11, true);
  addSectionTitle(slide, "قطاعات نخدمها", "Sectors we serve", { dark: true });
  addDivider(slide, true);
  const sectors = [
    ["الفعاليات والمعارض", "EVENTS & EXHIBITIONS"],
    ["الاجتماعات والمؤتمرات", "MEETINGS & CONFERENCES"],
    ["الضيافة والأغذية", "HOSPITALITY & FOOD"],
    ["الاحتياجات التشغيلية والصناعية", "OPERATIONS & INDUSTRY"]
  ];
  sectors.forEach((s, i) => {
    const x = 0.72 + i * 3.0;
    slide.addShape(SH.RECTANGLE, { x, y: 3.02, w: 2.7, h: 2.3, fill: { color: C.navyDark, transparency: 8 }, line: { color: C.teal, transparency: 48, width: 1.3 } });
    addEn(slide, String(i + 1).padStart(2, "0"), { x: x + 0.25, y: 3.3, w: 0.45, h: 0.26, fontSize: 13, bold: true, color: C.teal });
    addAr(slide, s[0], { x: x + 0.25, y: 4.0, w: 2.2, h: 0.6, fontSize: 16.2, bold: true, color: C.white, valign: "top" });
    addEn(slide, s[1], { x: x + 0.25, y: 4.78, w: 2.2, h: 0.32, fontSize: 7.8, bold: true, color: C.teal, charSpacing: 0.45, valign: "top" });
  });
  addBottomTag(slide, "حيث تكون الدقة في الموعد جزءًا من جودة المنتج.", "Where being on time is part of product quality.", true);
}

// 12 — Operating principles
{
  const slide = pptx.addSlide();
  slide.background = { color: C.white };
  addSlideChrome(slide, 12, false);
  addSectionTitle(slide, "مبادئ التشغيل", "Operating principles");
  addDivider(slide);
  const principles = [
    ["01", "الشفافية", "TRANSPARENCY", "توضيح الخيارات والتكلفة والوقت قبل تأكيد الطلب."],
    ["02", "المرونة", "FLEXIBILITY", "شبكة موردين تساعدنا على التعامل مع احتياجات متغيرة ومتنوعة."],
    ["03", "المتابعة", "FOLLOW-THROUGH", "تحديثات واضحة ومراقبة للتنفيذ حتى نقطة التسليم."]
  ];
  principles.forEach((p, i) => {
    const x = 0.72 + i * 4.05;
    slide.addShape(SH.RECTANGLE, { x, y: 2.55, w: 3.72, h: 3.55, fill: { color: i === 1 ? C.tealPale : C.paper }, line: { color: i === 1 ? C.teal : C.line, width: i === 1 ? 2 : 1 } });
    addEn(slide, p[0], { x: x + 0.3, y: 2.9, w: 0.5, h: 0.3, fontSize: 16, bold: true, color: C.tealDark });
    addAr(slide, p[1], { x: x + 0.3, y: 3.65, w: 3.12, h: 0.5, fontSize: 23, bold: true, color: C.ink });
    addEn(slide, p[2], { x: x + 0.3, y: 4.35, w: 3.12, h: 0.22, fontSize: 8.5, bold: true, color: C.tealDark, charSpacing: 0.7 });
    addAr(slide, p[3], { x: x + 0.3, y: 4.93, w: 3.12, h: 0.72, fontSize: 13.4, color: "435768", valign: "top" });
  });
  addBottomTag(slide, "السرعة مهمة. الوضوح أهم.", "Speed matters. Clarity matters more.");
}

// 13 — Contact
{
  const slide = pptx.addSlide();
  slide.background = { color: C.teal };
  addBrandPattern(slide, { x: 8.7, y: 0, w: 4.63, h: 7.5, transparency: 96 });
  slide.addShape(SH.RECTANGLE, { x: 0.55, y: 0.35, w: 3.08, h: 1.18, fill: { color: C.navyDark }, line: { color: C.navyDark } });
  addLogo(slide, true, 0.77, 0.53, 2.62);
  addAr(slide, "لديك طلب لا يحتمل الانتظار؟", { x: 6.15, y: 1.85, w: 6.05, h: 0.74, fontSize: 29, bold: true, color: C.navyDark });
  addEn(slide, "HAVE A REQUEST THAT CANNOT WAIT?", { x: 0.75, y: 1.97, w: 5.0, h: 0.35, fontSize: 13, bold: true, color: C.navyDark, charSpacing: 0.8 });
  addAr(slide, "أرسل المنتج، الكمية، موقع التسليم والموعد المطلوب. سنبدأ من الوقت المتاح.", { x: 6.15, y: 3.0, w: 6.05, h: 0.95, fontSize: 18, color: "153C48", valign: "top" });
  addEn(slide, "Share the item, quantity, delivery location and required time. We will start with the time available.", { x: 0.75, y: 3.0, w: 5.05, h: 0.82, fontSize: 14, color: "153C48", valign: "top" });
  slide.addShape(SH.LINE, { x: 0.75, y: 4.5, w: 11.45, h: 0, line: { color: C.navyDark, transparency: 35, width: 1 } });
  addAr(slide, "عبد العزيز الزعاقي", { x: 8.5, y: 4.85, w: 3.7, h: 0.4, fontSize: 18, bold: true, color: C.navyDark });
  addEn(slide, "CHIEF EXECUTIVE OFFICER", { x: 0.75, y: 4.91, w: 2.65, h: 0.25, fontSize: 9.5, bold: true, color: C.navyDark, charSpacing: 0.65 });
  addEn(slide, "+966 55 045 4848", { x: 0.75, y: 5.58, w: 2.6, h: 0.32, fontSize: 14, bold: true, color: C.navyDark });
  addEn(slide, "A.alzaaqi@LastMinute.com", { x: 3.75, y: 5.58, w: 3.4, h: 0.32, fontSize: 14, bold: true, color: C.navyDark });
  addAr(slide, "وساطة توريد مرنة للفعاليات والشركات", { x: 7.4, y: 6.62, w: 4.8, h: 0.26, fontSize: 11.5, bold: true, color: C.navyDark });
  addEn(slide, "LAST MINUTE", { x: 0.75, y: 6.68, w: 2.0, h: 0.22, fontSize: 9.5, bold: true, color: C.navyDark, charSpacing: 1.3 });
}

pptx.writeFile({ fileName: out });
