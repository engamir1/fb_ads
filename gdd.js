const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat,
  TableOfContents, ExternalHyperlink
} = require('docx');
const fs = require('fs');

const BLUE_DARK   = "0C3C78";
const BLUE_MID    = "1877F2";
const BLUE_LIGHT  = "D0E4FF";
const BLUE_BG     = "EBF3FF";
const ACCENT      = "42A5F5";
const GREEN       = "1B6B3A";
const GREEN_BG    = "D4EDDA";
const ORANGE      = "7B3F00";
const ORANGE_BG   = "FFF3CD";
const RED         = "7B0000";
const RED_BG      = "FDECEA";
const GRAY_BG     = "F5F5F5";
const GRAY_BORDER = "CCCCCC";
const WHITE       = "FFFFFF";
const TEXT_DARK   = "1A1A2E";
const TEXT_MID    = "444455";

const border = { style: BorderStyle.SINGLE, size: 1, color: GRAY_BORDER };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const W = 9360;
const cellM = { top: 100, bottom: 100, left: 140, right: 140 };

function hdrCell(text, w, fill=BLUE_DARK) {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: cellM,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: WHITE, size: 20, font: "Arial" })]
    })]
  });
}

function dataCell(text, w, fill=WHITE, bold=false, color=TEXT_DARK, align=AlignmentType.LEFT) {
  return new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: { fill, type: ShadingType.CLEAR },
    margins: cellM,
    children: [new Paragraph({
      alignment: align,
      children: [new TextRun({ text: String(text), bold, color, size: 18, font: "Arial" })]
    })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE_MID, space: 6 } },
    children: [new TextRun({ text, bold: true, size: 36, color: BLUE_DARK, font: "Arial" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, color: BLUE_MID, font: "Arial" })]
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: TEXT_DARK, font: "Arial" })]
  });
}

function h4(text) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 22, color: BLUE_DARK, font: "Arial", underline: {} })]
  });
}

function body(text, spaceBefore=60, spaceAfter=60) {
  return new Paragraph({
    spacing: { before: spaceBefore, after: spaceAfter },
    children: [new TextRun({ text, size: 20, color: TEXT_DARK, font: "Arial" })]
  });
}

function bullet(text, ref="bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, color: TEXT_DARK, font: "Arial" })]
  });
}

function subbullet(text) {
  return new Paragraph({
    numbering: { reference: "subbullets", level: 1 },
    spacing: { before: 30, after: 30 },
    children: [new TextRun({ text, size: 18, color: TEXT_MID, font: "Arial" })]
  });
}

function nb(text, ref="numbered") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 20, color: TEXT_DARK, font: "Arial" })]
  });
}

function callout(text, fillColor=BLUE_BG, textColor=BLUE_DARK, label="") {
  return new Paragraph({
    spacing: { before: 120, after: 120 },
    indent: { left: 360, right: 360 },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: BLUE_MID, space: 8 } },
    shading: { fill: fillColor, type: ShadingType.CLEAR },
    children: [
      label ? new TextRun({ text: label + " ", bold: true, size: 20, color: textColor, font: "Arial" }) : new TextRun(""),
      new TextRun({ text, size: 20, color: textColor, font: "Arial" })
    ].filter(r => r.text !== "")
  });
}

function divider() {
  return new Paragraph({
    spacing: { before: 200, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: GRAY_BORDER } },
    children: []
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function empty(n=1) {
  return Array.from({length:n}, () => new Paragraph({ children: [new TextRun("")] }));
}

function sectionBanner(text) {
  return new Paragraph({
    spacing: { before: 240, after: 200 },
    shading: { fill: BLUE_DARK, type: ShadingType.CLEAR },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text, bold: true, size: 32, color: WHITE, font: "Arial" })]
  });
}

// ============================================================
// COVER PAGE
// ============================================================
const coverChildren = [
  ...empty(4),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    children: [new TextRun({ text: "FACEBOOK ADS", bold: true, size: 72, color: BLUE_DARK, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({ text: "MULTI-PLATFORM MANAGER", bold: true, size: 56, color: BLUE_MID, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 80 },
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: BLUE_MID }, bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE_MID } },
    shading: { fill: BLUE_BG, type: ShadingType.CLEAR },
    children: [new TextRun({ text: "Game Design Document  ·  GDD v1.0", bold: false, size: 28, color: TEXT_MID, font: "Arial" })]
  }),
  ...empty(2),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 60 },
    children: [new TextRun({ text: "وثيقة التصميم الشاملة", bold: true, size: 36, color: BLUE_DARK, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({ text: "الإصدار الأول · مايو 2026", size: 22, color: TEXT_MID, font: "Arial" })]
  }),
  ...empty(3),
  new Table({
    width: { size: 6000, type: WidthType.DXA },
    columnWidths: [2400, 3600],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, width:{size:2400,type:WidthType.DXA}, shading:{fill:BLUE_DARK,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({alignment:AlignmentType.RIGHT, children:[new TextRun({text:"المشروع", bold:true, size:20, color:WHITE, font:"Arial"})]})] }),
        new TableCell({ borders, width:{size:3600,type:WidthType.DXA}, shading:{fill:BLUE_BG,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({children:[new TextRun({text:"Facebook Ads Multi-Platform Manager", size:20, color:TEXT_DARK, font:"Arial"})]})] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width:{size:2400,type:WidthType.DXA}, shading:{fill:BLUE_DARK,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({alignment:AlignmentType.RIGHT, children:[new TextRun({text:"الإصدار", bold:true, size:20, color:WHITE, font:"Arial"})]})] }),
        new TableCell({ borders, width:{size:3600,type:WidthType.DXA}, shading:{fill:WHITE,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({children:[new TextRun({text:"v1.0 — Initial Release", size:20, color:TEXT_DARK, font:"Arial"})]})] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width:{size:2400,type:WidthType.DXA}, shading:{fill:BLUE_DARK,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({alignment:AlignmentType.RIGHT, children:[new TextRun({text:"التاريخ", bold:true, size:20, color:WHITE, font:"Arial"})]})] }),
        new TableCell({ borders, width:{size:3600,type:WidthType.DXA}, shading:{fill:BLUE_BG,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({children:[new TextRun({text:"مايو 2026", size:20, color:TEXT_DARK, font:"Arial"})]})] })
      ]}),
      new TableRow({ children: [
        new TableCell({ borders, width:{size:2400,type:WidthType.DXA}, shading:{fill:BLUE_DARK,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({alignment:AlignmentType.RIGHT, children:[new TextRun({text:"نوع الوثيقة", bold:true, size:20, color:WHITE, font:"Arial"})]})] }),
        new TableCell({ borders, width:{size:3600,type:WidthType.DXA}, shading:{fill:WHITE,type:ShadingType.CLEAR}, margins:cellM,
          children:[new Paragraph({children:[new TextRun({text:"Game Design Document (GDD)", size:20, color:TEXT_DARK, font:"Arial"})]})] })
      ]}),
    ]
  }),
];

// ============================================================
// SECTION 1 - EXECUTIVE SUMMARY
// ============================================================
const sec1 = [
  pageBreak(),
  sectionBanner("القسم الأول: الملخص التنفيذي"),
  h1("1. نظرة عامة على المشروع"),
  h2("1.1 رؤية المنتج"),
  body("Facebook Ads Multi-Platform Manager هو نظام إدارة إعلانات متكامل يهدف إلى تمكين المعلنين من التحكم الكامل في حملاتهم الإعلانية عبر واجهة واحدة شاملة، تربط بين قوة Facebook Marketing API وسهولة الاستخدام اليومي. يهدف النظام إلى إزالة الحاجة للتبديل بين واجهات متعددة وتوفير رؤية 360 درجة لكل أنشطة الإعلان."),
  body("الهدف النهائي: تطبيق يعمل على جميع المنصات (Web, Desktop, Mobile, CLI) ويتيح التحكم الكامل في كل مستوى من مستويات هرم الإعلانات من الحساب الإعلاني حتى أصغر Creative Asset."),

  h2("1.2 المشكلة التي يحلها"),
  bullet("تشتت بيانات الحملات عبر حسابات متعددة بدون مركزية"),
  bullet("صعوبة مقارنة الأداء بين الحملات والمجموعات الإعلانية في لحظة واحدة"),
  bullet("غياب أدوات Bulk Actions فعّالة لتعديل حملات كثيرة دفعة واحدة"),
  bullet("ضعف الرؤية التحليلية العميقة في واجهة Ads Manager الافتراضية"),
  bullet("صعوبة إدارة الإعلانات من الأجهزة المحمولة"),
  bullet("غياب أتمتة القواعد والتنبيهات الذكية"),

  h2("1.3 الجمهور المستهدف"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2340, 3510, 3510],
    rows: [
      new TableRow({ children: [hdrCell("الشريحة", 2340), hdrCell("الوصف", 3510), hdrCell("الاحتياج الأساسي", 3510)] }),
      new TableRow({ children: [dataCell("Media Buyers", 2340, GRAY_BG, true), dataCell("محترفو شراء الإعلانات يديرون حسابات متعددة بميزانيات كبيرة", 3510), dataCell("Bulk actions، automation، reporting متقدم", 3510)] }),
      new TableRow({ children: [dataCell("أصحاب التجارة الإلكترونية", 2340, WHITE, true), dataCell("مالكو متاجر يديرون إعلاناتهم بأنفسهم", 3510), dataCell("سهولة الاستخدام، واجهة بسيطة، تقارير ROAS", 3510)] }),
      new TableRow({ children: [dataCell("Agencies", 2340, GRAY_BG, true), dataCell("وكالات تدير حسابات عملاء متعددة", 3510), dataCell("Multi-account، white-label، client reports", 3510)] }),
      new TableRow({ children: [dataCell("المسوقون المستقلون", 2340, WHITE, true), dataCell("Freelancers يخدمون عملاء متعددين", 3510, ), dataCell("إدارة من موبايل، تقارير سريعة للعملاء", 3510)] }),
    ]
  }),
];

// ============================================================
// SECTION 2 - FACEBOOK API: ALL DATA AVAILABLE
// ============================================================
const sec2 = [
  pageBreak(),
  sectionBanner("القسم الثاني: بيانات Facebook API — التغطية الكاملة"),
  h1("2. هرم بيانات Facebook Marketing API"),
  body("يعمل Facebook Marketing API وفق هيكل هرمي من 5 مستويات. كل مستوى يرث إعدادات ما فوقه ويمكن التحكم به بشكل مستقل. فيما يلي التوثيق الكامل لكل حقل متاح في كل مستوى."),

  h2("2.1 المستوى الأول: Ad Account (الحساب الإعلاني)"),
  h3("2.1.1 الحقول الأساسية المتاحة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2600, 2200, 1800, 2760],
    rows: [
      new TableRow({ children: [hdrCell("الحقل (Field)", 2600), hdrCell("النوع", 2200), hdrCell("قابل للتعديل", 1800), hdrCell("الوصف", 2760)] }),
      new TableRow({ children: [dataCell("id", 2600, GRAY_BG), dataCell("string", 2200, GRAY_BG), dataCell("لا", 1800, GRAY_BG, false, RED), dataCell("معرف الحساب الفريد (act_XXXXXXX)", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 2600), dataCell("string", 2200), dataCell("نعم", 1800, WHITE, false, GREEN), dataCell("اسم الحساب الإعلاني", 2760)] }),
      new TableRow({ children: [dataCell("account_status", 2600, GRAY_BG), dataCell("enum int", 2200, GRAY_BG), dataCell("Admin فقط", 1800, GRAY_BG, false, ORANGE), dataCell("1=نشط, 2=معطل, 3=غير مفعل, 7=ملغي, 9=منتهي, 101=بانتظار, 201=محدود", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("amount_spent", 2600), dataCell("string (cents)", 2200), dataCell("لا", 1800, WHITE, false, RED), dataCell("إجمالي الإنفاق بالسنتات منذ الإنشاء", 2760)] }),
      new TableRow({ children: [dataCell("balance", 2600, GRAY_BG), dataCell("string (cents)", 2200, GRAY_BG), dataCell("لا", 1800, GRAY_BG, false, RED), dataCell("الرصيد الحالي للحساب", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("currency", 2600), dataCell("string ISO-4217", 2200), dataCell("نعم", 1800, WHITE, false, GREEN), dataCell("عملة الحساب (USD, EGP, SAR...)", 2760)] }),
      new TableRow({ children: [dataCell("timezone_name", 2600, GRAY_BG), dataCell("string", 2200, GRAY_BG), dataCell("نعم", 1800, GRAY_BG, false, GREEN), dataCell("المنطقة الزمنية (Africa/Cairo, UTC...)", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("created_time", 2600), dataCell("datetime", 2200), dataCell("لا", 1800, WHITE, false, RED), dataCell("تاريخ إنشاء الحساب", 2760)] }),
      new TableRow({ children: [dataCell("business", 2600, GRAY_BG), dataCell("object", 2200, GRAY_BG), dataCell("لا", 1800, GRAY_BG, false, RED), dataCell("البيزنس مانجر المرتبط (id, name)", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("owner", 2600), dataCell("object", 2200), dataCell("لا", 1800, WHITE, false, RED), dataCell("مالك الحساب (id, name)", 2760)] }),
      new TableRow({ children: [dataCell("spend_cap", 2600, GRAY_BG), dataCell("string (cents)", 2200, GRAY_BG), dataCell("نعم", 1800, GRAY_BG, false, GREEN), dataCell("الحد الأقصى للإنفاق الكلي على الحساب", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("funding_source", 2600), dataCell("string", 2200), dataCell("لا", 1800, WHITE, false, RED), dataCell("معرف مصدر التمويل المرتبط", 2760)] }),
      new TableRow({ children: [dataCell("funding_source_details", 2600, GRAY_BG), dataCell("object", 2200, GRAY_BG), dataCell("لا", 1800, GRAY_BG, false, RED), dataCell("تفاصيل مصدر الدفع (نوع البطاقة، آخر 4 أرقام...)", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("age", 2600), dataCell("float (days)", 2200), dataCell("لا", 1800, WHITE, false, RED), dataCell("عمر الحساب بالأيام", 2760)] }),
      new TableRow({ children: [dataCell("brand_safety_content_filter_levels", 2600, GRAY_BG), dataCell("array enum", 2200, GRAY_BG), dataCell("نعم", 1800, GRAY_BG, false, GREEN), dataCell("مستويات فلترة Brand Safety", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("attribution_spec", 2600), dataCell("array", 2200), dataCell("نعم", 1800, WHITE, false, GREEN), dataCell("نوافذ الإسناد الافتراضية للحساب", 2760)] }),
      new TableRow({ children: [dataCell("capabilities", 2600, GRAY_BG), dataCell("array", 2200, GRAY_BG), dataCell("لا", 1800, GRAY_BG, false, RED), dataCell("الإمكانيات المتاحة (creative_hub, etc.)", 2760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("is_notifications_enabled", 2600), dataCell("bool", 2200), dataCell("نعم", 1800, WHITE, false, GREEN), dataCell("هل التنبيهات مفعلة", 2760)] }),
      new TableRow({ children: [dataCell("min_campaign_group_spend_cap", 2600, GRAY_BG), dataCell("int", 2200, GRAY_BG), dataCell("لا", 1800, GRAY_BG, false, RED), dataCell("الحد الأدنى لـ spend cap على مستوى الحملة", 2760, GRAY_BG)] }),
    ]
  }),

  h2("2.2 المستوى الثاني: Campaign (الحملة الإعلانية)"),
  h3("2.2.1 الحقول الكاملة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2500, 2100, 1600, 3160],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2500), hdrCell("النوع", 2100), hdrCell("قابل للتعديل", 1600), hdrCell("القيم المتاحة / الوصف", 3160)] }),
      new TableRow({ children: [dataCell("id", 2500, GRAY_BG), dataCell("string", 2100, GRAY_BG), dataCell("لا", 1600, GRAY_BG, false, RED), dataCell("معرف الحملة الفريد", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 2500), dataCell("string", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("اسم الحملة", 3160)] }),
      new TableRow({ children: [dataCell("objective", 2500, GRAY_BG), dataCell("enum", 2100, GRAY_BG), dataCell("لا (بعد الإنشاء)", 1600, GRAY_BG, false, ORANGE), dataCell("CONVERSIONS, LEAD_GENERATION, LINK_CLICKS, REACH, VIDEO_VIEWS, POST_ENGAGEMENT, BRAND_AWARENESS, APP_INSTALLS, MESSAGES, PAGE_LIKES, STORE_VISITS, PRODUCT_CATALOG_SALES, EVENT_RESPONSES", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("status", 2500), dataCell("enum", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("ACTIVE, PAUSED, DELETED, ARCHIVED", 3160)] }),
      new TableRow({ children: [dataCell("effective_status", 2500, GRAY_BG), dataCell("enum", 2100, GRAY_BG), dataCell("لا (محسوب)", 1600, GRAY_BG, false, RED), dataCell("الحالة الفعلية مع الأخذ بعين الاعتبار حالة الحساب والفواتير", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("configured_status", 2500), dataCell("enum", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("ACTIVE, PAUSED, DELETED, ARCHIVED — الحالة المضبوطة من المستخدم", 3160)] }),
      new TableRow({ children: [dataCell("daily_budget", 2500, GRAY_BG), dataCell("string (cents)", 2100, GRAY_BG), dataCell("نعم", 1600, GRAY_BG, false, GREEN), dataCell("الميزانية اليومية بالسنتات (يستبعد lifetime_budget)", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("lifetime_budget", 2500), dataCell("string (cents)", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("الميزانية الإجمالية بالسنتات (يستبعد daily_budget)", 3160)] }),
      new TableRow({ children: [dataCell("budget_remaining", 2500, GRAY_BG), dataCell("string (cents)", 2100, GRAY_BG), dataCell("لا (محسوب)", 1600, GRAY_BG, false, RED), dataCell("الميزانية المتبقية الحالية", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("spend_cap", 2500), dataCell("string (cents)", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("الحد الأقصى للإنفاق على مستوى الحملة", 3160)] }),
      new TableRow({ children: [dataCell("bid_strategy", 2500, GRAY_BG), dataCell("enum", 2100, GRAY_BG), dataCell("نعم", 1600, GRAY_BG, false, GREEN), dataCell("LOWEST_COST_WITHOUT_CAP, LOWEST_COST_WITH_BID_CAP, COST_CAP, LOWEST_COST_WITH_MIN_ROAS", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("buying_type", 2500), dataCell("enum", 2100), dataCell("لا (بعد الإنشاء)", 1600, WHITE, false, ORANGE), dataCell("AUCTION, RESERVED, DEPRECATED_REACH_BLOCK", 3160)] }),
      new TableRow({ children: [dataCell("start_time", 2500, GRAY_BG), dataCell("datetime", 2100, GRAY_BG), dataCell("نعم", 1600, GRAY_BG, false, GREEN), dataCell("تاريخ ووقت بدء الحملة (ISO 8601)", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("stop_time", 2500), dataCell("datetime", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("تاريخ ووقت انتهاء الحملة (ISO 8601)", 3160)] }),
      new TableRow({ children: [dataCell("created_time", 2500, GRAY_BG), dataCell("datetime", 2100, GRAY_BG), dataCell("لا", 1600, GRAY_BG, false, RED), dataCell("تاريخ إنشاء الحملة", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("updated_time", 2500), dataCell("datetime", 2100), dataCell("لا", 1600, WHITE, false, RED), dataCell("آخر تعديل على الحملة", 3160)] }),
      new TableRow({ children: [dataCell("special_ad_categories", 2500, GRAY_BG), dataCell("array enum", 2100, GRAY_BG), dataCell("نعم", 1600, GRAY_BG, false, GREEN), dataCell("EMPLOYMENT, HOUSING, CREDIT, ISSUES_ELECTIONS_POLITICS, NONE", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("pacing_type", 2500), dataCell("array", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("[standard] أو [no_pacing] — توزيع الإنفاق عبر اليوم", 3160)] }),
      new TableRow({ children: [dataCell("promoted_object", 2500, GRAY_BG), dataCell("object", 2100, GRAY_BG), dataCell("لا (بعد الإنشاء)", 1600, GRAY_BG, false, ORANGE), dataCell("pixel_id, page_id, app_id, product_catalog_id — الكائن المُعلن عنه", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("budget_rebalance_flag", 2500), dataCell("bool", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("إعادة توزيع الميزانية تلقائياً بين الأد سيت", 3160)] }),
      new TableRow({ children: [dataCell("is_skadnetwork_attribution", 2500, GRAY_BG), dataCell("bool", 2100, GRAY_BG), dataCell("نعم", 1600, GRAY_BG, false, GREEN), dataCell("تفعيل SKAdNetwork لحملات iOS 14.5+", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("adlabels", 2500), dataCell("array", 2100), dataCell("نعم", 1600, WHITE, false, GREEN), dataCell("تسميات التصنيف للحملة", 3160)] }),
      new TableRow({ children: [dataCell("recommendations", 2500, GRAY_BG), dataCell("array", 2100, GRAY_BG), dataCell("لا", 1600, GRAY_BG, false, RED), dataCell("توصيات Facebook لتحسين الحملة", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("source_campaign_id", 2500), dataCell("string", 2100), dataCell("لا", 1600, WHITE, false, RED), dataCell("ID الحملة الأصلية إذا كانت منسوخة", 3160)] }),
      new TableRow({ children: [dataCell("issues_info", 2500, GRAY_BG), dataCell("array", 2100, GRAY_BG), dataCell("لا", 1600, GRAY_BG, false, RED), dataCell("قائمة المشاكل وأكوادها وأسبابها", 3160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("topline_id", 2500), dataCell("string", 2100), dataCell("لا", 1600, WHITE, false, RED), dataCell("معرف الـ Topline (للإعلانات المحجوزة)", 3160)] }),
    ]
  }),

  h2("2.3 المستوى الثالث: Ad Set (المجموعة الإعلانية)"),
  body("الأد سيت هو المستوى الأكثر ثراءً بالبيانات في الهرم، إذ يحتوي على كل إعدادات الاستهداف والجدولة والميزانية والأوبتيمايزيشن."),
  h3("2.3.1 الحقول الأساسية والميزانية"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2500, 2000, 1500, 3360],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2500), hdrCell("النوع", 2000), hdrCell("قابل للتعديل", 1500), hdrCell("الوصف", 3360)] }),
      new TableRow({ children: [dataCell("id", 2500, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("لا", 1500, GRAY_BG, false, RED), dataCell("معرف المجموعة الإعلانية", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 2500), dataCell("string", 2000), dataCell("نعم", 1500, WHITE, false, GREEN), dataCell("اسم المجموعة الإعلانية", 3360)] }),
      new TableRow({ children: [dataCell("campaign_id", 2500, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("لا", 1500, GRAY_BG, false, RED), dataCell("معرف الحملة الأم", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("daily_budget", 2500), dataCell("string (cents)", 2000), dataCell("نعم", 1500, WHITE, false, GREEN), dataCell("الميزانية اليومية على مستوى الأد سيت", 3360)] }),
      new TableRow({ children: [dataCell("lifetime_budget", 2500, GRAY_BG), dataCell("string (cents)", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("الميزانية الإجمالية على مستوى الأد سيت", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("budget_remaining", 2500), dataCell("string (cents)", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("الميزانية المتبقية الحالية", 3360)] }),
      new TableRow({ children: [dataCell("daily_spend_cap", 2500, GRAY_BG), dataCell("string (cents)", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("حد أقصى للإنفاق اليومي (لا يوقف الأد سيت)", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("lifetime_spend_cap", 2500), dataCell("string (cents)", 2000), dataCell("نعم", 1500, WHITE, false, GREEN), dataCell("حد أقصى للإنفاق الإجمالي", 3360)] }),
      new TableRow({ children: [dataCell("daily_min_spend_target", 2500, GRAY_BG), dataCell("string (cents)", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("الحد الأدنى للإنفاق اليومي (للحملات المحجوزة)", 3360, GRAY_BG)] }),
    ]
  }),

  h3("2.3.2 حقول الاستهداف (Targeting)"),
  callout("حقل targeting هو كائن JSON معقد يحتوي على عشرات الخصائص. فيما يلي التوثيق الكامل لكل خصائصه.", BLUE_BG, BLUE_DARK, "ملاحظة:"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2800, 2000, 4560],
    rows: [
      new TableRow({ children: [hdrCell("الخاصية", 2800), hdrCell("النوع", 2000), hdrCell("الوصف والقيم", 4560)] }),
      new TableRow({ children: [dataCell("age_min", 2800, GRAY_BG), dataCell("int", 2000, GRAY_BG), dataCell("الحد الأدنى للعمر (13 - 65+)", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("age_max", 2800), dataCell("int", 2000), dataCell("الحد الأقصى للعمر (13 - 65)", 4560)] }),
      new TableRow({ children: [dataCell("genders", 2800, GRAY_BG), dataCell("array int", 2000, GRAY_BG), dataCell("[1]=ذكور, [2]=إناث, [1,2]=الجميع", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("geo_locations", 2800), dataCell("object", 2000), dataCell("countries: ['EG','SA'], regions, cities, zips, location_types: ['home','recent','travel_in']", 4560)] }),
      new TableRow({ children: [dataCell("excluded_geo_locations", 2800, GRAY_BG), dataCell("object", 2000, GRAY_BG), dataCell("مواقع مستثناة — نفس هيكل geo_locations", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("locales", 2800), dataCell("array int", 2000), dataCell("اللغات المستهدفة (6=Arabic, 24=English...)", 4560)] }),
      new TableRow({ children: [dataCell("interests", 2800, GRAY_BG), dataCell("array object", 2000, GRAY_BG), dataCell("[{id, name}] — اهتمامات المستخدمين من فيسبوك", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("behaviors", 2800), dataCell("array object", 2000), dataCell("[{id, name}] — سلوكيات الشراء والسفر والتقنية وغيرها", 4560)] }),
      new TableRow({ children: [dataCell("demographics", 2800, GRAY_BG), dataCell("array object", 2000, GRAY_BG), dataCell("[{id, name}] — التعليم، الوظيفة، الدخل، الحياة الأسرية", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("connections", 2800), dataCell("array object", 2000), dataCell("من أعجب بصفحة أو استخدم تطبيقاً", 4560)] }),
      new TableRow({ children: [dataCell("excluded_connections", 2800, GRAY_BG), dataCell("array object", 2000, GRAY_BG), dataCell("استثناء من أعجب بالصفحة", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("friends_of_connections", 2800), dataCell("array object", 2000), dataCell("أصدقاء من تربطهم علاقة بصفحة أو تطبيق", 4560)] }),
      new TableRow({ children: [dataCell("custom_audiences", 2800, GRAY_BG), dataCell("array object", 2000, GRAY_BG), dataCell("[{id}] — Custom Audiences و Lookalike Audiences", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("excluded_custom_audiences", 2800), dataCell("array object", 2000), dataCell("[{id}] — Custom Audiences مستثناة", 4560)] }),
      new TableRow({ children: [dataCell("flexible_spec", 2800, GRAY_BG), dataCell("array object", 2000, GRAY_BG), dataCell("Detailed Targeting Expansion — مجموعات OR بين المعايير", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("exclusions", 2800), dataCell("object", 2000), dataCell("نفس هيكل flexible_spec للاستثناءات", 4560)] }),
      new TableRow({ children: [dataCell("publisher_platforms", 2800, GRAY_BG), dataCell("array string", 2000, GRAY_BG), dataCell("facebook, instagram, audience_network, messenger", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("facebook_positions", 2800), dataCell("array string", 2000), dataCell("feed, right_hand_column, instant_article, marketplace, video_feeds, story, search, instream_video, reels_overlay", 4560)] }),
      new TableRow({ children: [dataCell("instagram_positions", 2800, GRAY_BG), dataCell("array string", 2000, GRAY_BG), dataCell("stream, story, explore, reels, explore_home, profile_feed, ig_search", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("audience_network_positions", 2800), dataCell("array string", 2000), dataCell("classic, instream_video, rewarded_video", 4560)] }),
      new TableRow({ children: [dataCell("messenger_positions", 2800, GRAY_BG), dataCell("array string", 2000, GRAY_BG), dataCell("messenger_home, story, sponsored_messages", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("device_platforms", 2800), dataCell("array string", 2000), dataCell("mobile, desktop", 4560)] }),
      new TableRow({ children: [dataCell("user_os", 2800, GRAY_BG), dataCell("array string", 2000, GRAY_BG), dataCell("iOS, Android, iOS_ver_X_0_to_X_9 — نظام التشغيل المستهدف", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("user_device", 2800), dataCell("array string", 2000), dataCell("iPhone, iPad, Android_Smartphone — نوع الجهاز", 4560)] }),
      new TableRow({ children: [dataCell("wireless_carrier", 2800, GRAY_BG), dataCell("array string", 2000, GRAY_BG), dataCell("Wifi — الاستهداف بحسب نوع الاتصال", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("targeting_optimization", 2800), dataCell("string", 2000), dataCell("expansion_all — السماح لـ Facebook بتوسيع الاستهداف", 4560)] }),
    ]
  }),

  h3("2.3.3 حقول الجدولة والأوبتيمايزيشن"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2800, 2000, 4560],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2800), hdrCell("النوع", 2000), hdrCell("الوصف", 4560)] }),
      new TableRow({ children: [dataCell("optimization_goal", 2800, GRAY_BG), dataCell("enum", 2000, GRAY_BG), dataCell("APP_INSTALLS, BRAND_AWARENESS, CLICKS, ENGAGED_USERS, EVENT_RESPONSES, IMPRESSIONS, LEAD_GENERATION, LINK_CLICKS, NONE, OFFSITE_CONVERSIONS, PAGE_ENGAGEMENT, PAGE_LIKES, POST_ENGAGEMENT, QUALITY_LEAD, REACH, REPLIES, THRUPLAY, VALUE, VISIT_INSTAGRAM_PROFILE", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("billing_event", 2800), dataCell("enum", 2000), dataCell("APP_INSTALLS, CLICKS, IMPRESSIONS, LINK_CLICKS, NONE, OFFER_CLAIMS, PAGE_LIKES, POST_ENGAGEMENT, THRUPLAY, PURCHASE — حدث الفوترة", 4560)] }),
      new TableRow({ children: [dataCell("bid_amount", 2800, GRAY_BG), dataCell("int (cents)", 2000, GRAY_BG), dataCell("المبلغ الأقصى للمزايدة في حال bid_strategy = BID_CAP أو COST_CAP", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("bid_strategy", 2800), dataCell("enum", 2000), dataCell("LOWEST_COST_WITHOUT_CAP, LOWEST_COST_WITH_BID_CAP, COST_CAP", 4560)] }),
      new TableRow({ children: [dataCell("start_time", 2800, GRAY_BG), dataCell("datetime", 2000, GRAY_BG), dataCell("وقت بدء الأد سيت", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("end_time", 2800), dataCell("datetime", 2000), dataCell("وقت انتهاء الأد سيت", 4560)] }),
      new TableRow({ children: [dataCell("pacing_type", 2800, GRAY_BG), dataCell("array string", 2000, GRAY_BG), dataCell("[standard] = توزيع متساوٍ, [no_pacing] = إنفاق أسرع", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("destination_type", 2800), dataCell("enum", 2000), dataCell("WEBSITE, APP, MESSENGER, APPLINKS_AUTOMATIC, FACEBOOK, INSTAGRAM_DIRECT, WHATSAPP", 4560)] }),
      new TableRow({ children: [dataCell("promoted_object", 2800, GRAY_BG), dataCell("object", 2000, GRAY_BG), dataCell("pixel_id, page_id, app_id, custom_conversion_id, pixel_rule — الكائن المُحسَّن له", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("rf_prediction_id", 2800), dataCell("string", 2000), dataCell("معرف توقع Reach & Frequency (للحملات المحجوزة)", 4560)] }),
      new TableRow({ children: [dataCell("frequency_control_specs", 2800, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{event: 'IMPRESSIONS', interval_days: 7, max_frequency: 3}] — تحديد تكرار ظهور الإعلان", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("time_based_ad_rotation_id_blocks", 2800), dataCell("array", 2000), dataCell("بلوكات دوران الإعلانات بناءً على الوقت", 4560)] }),
      new TableRow({ children: [dataCell("is_budget_schedule_enabled", 2800, GRAY_BG), dataCell("bool", 2000, GRAY_BG), dataCell("هل جدولة الميزانية مفعلة (CBO)", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("attribution_spec", 2800), dataCell("array", 2000), dataCell("[{event_type, window_days}] — نوافذ الإسناد المخصصة للأد سيت", 4560)] }),
      new TableRow({ children: [dataCell("use_new_app_click", 2800, GRAY_BG), dataCell("bool", 2000, GRAY_BG), dataCell("استخدام تتبع النقرات الجديد للتطبيقات", 4560, GRAY_BG)] }),
    ]
  }),

  h2("2.4 المستوى الرابع: Ad (الإعلان الفردي)"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2500, 2000, 1500, 3360],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2500), hdrCell("النوع", 2000), hdrCell("قابل للتعديل", 1500), hdrCell("الوصف", 3360)] }),
      new TableRow({ children: [dataCell("id", 2500, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("لا", 1500, GRAY_BG, false, RED), dataCell("معرف الإعلان", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 2500), dataCell("string", 2000), dataCell("نعم", 1500, WHITE, false, GREEN), dataCell("اسم الإعلان", 3360)] }),
      new TableRow({ children: [dataCell("adset_id", 2500, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("لا", 1500, GRAY_BG, false, RED), dataCell("معرف الأد سيت الأب", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("campaign_id", 2500), dataCell("string", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("معرف الحملة (للراحة)", 3360)] }),
      new TableRow({ children: [dataCell("status", 2500, GRAY_BG), dataCell("enum", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("ACTIVE, PAUSED, DELETED, ARCHIVED", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("effective_status", 2500), dataCell("enum", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("ACTIVE, PAUSED, DELETED, PENDING_REVIEW, DISAPPROVED, PREAPPROVED, PENDING_BILLING_INFO, CAMPAIGN_PAUSED, ARCHIVED, ADSET_PAUSED, IN_PROCESS, WITH_ISSUES", 3360)] }),
      new TableRow({ children: [dataCell("creative", 2500, GRAY_BG), dataCell("object", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("{id} — معرف Creative الإعلاني المرتبط", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("bid_amount", 2500), dataCell("int (cents)", 2000), dataCell("نعم", 1500, WHITE, false, GREEN), dataCell("المزايدة على مستوى الإعلان (يتجاوز مزايدة الأد سيت)", 3360)] }),
      new TableRow({ children: [dataCell("bid_type", 2500, GRAY_BG), dataCell("enum", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("CPC, CPM, MULTI_PREMIUM, ABSOLUTE_OCPM, CPA", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("conversion_domain", 2500), dataCell("string", 2000), dataCell("نعم", 1500, WHITE, false, GREEN), dataCell("النطاق المستخدم للتحويلات", 3360)] }),
      new TableRow({ children: [dataCell("tracking_specs", 2500, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("مصفوفة أكواد التتبع (pixel events, app events...)", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("preview_shareable_link", 2500), dataCell("string URL", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("رابط معاينة الإعلان كما يظهر للجمهور", 3360)] }),
      new TableRow({ children: [dataCell("issues_info", 2500, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("لا", 1500, GRAY_BG, false, RED), dataCell("[{error_code, error_summary, error_message, level, action_required}] — تفاصيل المشاكل", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("recommendations", 2500), dataCell("array", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("توصيات Facebook لتحسين الإعلان", 3360)] }),
      new TableRow({ children: [dataCell("adlabels", 2500, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("نعم", 1500, GRAY_BG, false, GREEN), dataCell("تسميات تصنيف الإعلان", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("source_ad_id", 2500), dataCell("string", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("معرف الإعلان الأصلي إذا كان منسوخاً", 3360)] }),
      new TableRow({ children: [dataCell("created_time", 2500, GRAY_BG), dataCell("datetime", 2000, GRAY_BG), dataCell("لا", 1500, GRAY_BG, false, RED), dataCell("تاريخ إنشاء الإعلان", 3360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("updated_time", 2500), dataCell("datetime", 2000), dataCell("لا", 1500, WHITE, false, RED), dataCell("آخر تعديل", 3360)] }),
    ]
  }),

  h2("2.5 المستوى الخامس: Ad Creative (المحتوى الإبداعي)"),
  h3("2.5.1 الحقول الأساسية"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2600, 2000, 4760],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2600), hdrCell("النوع", 2000), hdrCell("الوصف", 4760)] }),
      new TableRow({ children: [dataCell("id", 2600, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("معرف الـ Creative", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 2600), dataCell("string", 2000), dataCell("اسم الـ Creative", 4760)] }),
      new TableRow({ children: [dataCell("title", 2600, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("عنوان الإعلان (Headline)", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("body", 2600), dataCell("string", 2000), dataCell("نص الإعلان الرئيسي", 4760)] }),
      new TableRow({ children: [dataCell("link_url", 2600, GRAY_BG), dataCell("string URL", 2000, GRAY_BG), dataCell("الرابط المقصود", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("image_url", 2600), dataCell("string URL", 2000), dataCell("رابط الصورة", 4760)] }),
      new TableRow({ children: [dataCell("image_hash", 2600, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("هاش الصورة في مكتبة الصور", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("video_id", 2600), dataCell("string", 2000), dataCell("معرف الفيديو في Facebook Videos", 4760)] }),
      new TableRow({ children: [dataCell("call_to_action", 2600, GRAY_BG), dataCell("object", 2000, GRAY_BG), dataCell("{type: 'SHOP_NOW', value: {link, link_caption}} — زر الدعوة للعمل", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("object_story_spec", 2600), dataCell("object", 2000), dataCell("page_id + (link_data | video_data | photo_data | text_data | offer_data | template_data) — هيكل الإعلان الكامل", 4760)] }),
      new TableRow({ children: [dataCell("asset_feed_spec", 2600, GRAY_BG), dataCell("object", 2000, GRAY_BG), dataCell("Dynamic Creative — مصفوفات عناوين وصور وأزرار للاختبار التلقائي", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("instagram_actor_id", 2600), dataCell("string", 2000), dataCell("معرف حساب انستغرام المُرسِل للإعلان", 4760)] }),
      new TableRow({ children: [dataCell("use_page_actor_override", 2600, GRAY_BG), dataCell("bool", 2000, GRAY_BG), dataCell("هل يتم العرض بهوية الصفحة على انستغرام", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("effective_instagram_media_id", 2600), dataCell("string", 2000), dataCell("معرف البوست الفعلي على انستغرام", 4760)] }),
      new TableRow({ children: [dataCell("url_tags", 2600, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("UTM Parameters المضافة تلقائياً للروابط", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("template_url", 2600), dataCell("string URL", 2000), dataCell("قالب الرابط الديناميكي للـ DPA", 4760)] }),
      new TableRow({ children: [dataCell("product_set_id", 2600, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("مجموعة المنتجات المستخدمة في إعلانات الكتالوج", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("degrees_of_freedom_spec", 2600), dataCell("object", 2000), dataCell("خصائص Advantage+ Creative التلقائية", 4760)] }),
      new TableRow({ children: [dataCell("thumbnail_url", 2600, GRAY_BG), dataCell("string URL", 2000, GRAY_BG), dataCell("رابط صورة مصغرة للفيديو", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("status", 2600), dataCell("enum", 2000), dataCell("ACTIVE, DELETED, IN_PROCESS, WITH_ISSUES", 4760)] }),
    ]
  }),

  h2("2.6 Insights API — مؤشرات الأداء الكاملة"),
  body("Insights API هو أقوى نقطة بيانات في Facebook Marketing API. يتيح الحصول على مئات المؤشرات بمستويات تقسيم متعددة."),
  h3("2.6.1 المؤشرات الأساسية (Basic Metrics)"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2600, 2000, 4760],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2600), hdrCell("الوحدة", 2000), hdrCell("الوصف", 4760)] }),
      new TableRow({ children: [dataCell("impressions", 2600, GRAY_BG), dataCell("int", 2000, GRAY_BG), dataCell("عدد مرات ظهور الإعلان", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("reach", 2600), dataCell("int", 2000), dataCell("عدد الأشخاص الفريدين الذين رأوا الإعلان", 4760)] }),
      new TableRow({ children: [dataCell("frequency", 2600, GRAY_BG), dataCell("float", 2000, GRAY_BG), dataCell("متوسط عدد مرات رؤية الإعلان لكل شخص", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("clicks", 2600), dataCell("int", 2000), dataCell("إجمالي النقرات (جميع أنواع النقرات)", 4760)] }),
      new TableRow({ children: [dataCell("unique_clicks", 2600, GRAY_BG), dataCell("int", 2000, GRAY_BG), dataCell("النقرات من أشخاص فريدين", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("ctr", 2600), dataCell("percent", 2000), dataCell("Click-Through Rate = clicks / impressions × 100", 4760)] }),
      new TableRow({ children: [dataCell("unique_ctr", 2600, GRAY_BG), dataCell("percent", 2000, GRAY_BG), dataCell("CTR الفريد = unique_clicks / reach × 100", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("spend", 2600), dataCell("float ($)", 2000), dataCell("إجمالي الإنفاق في الفترة المحددة", 4760)] }),
      new TableRow({ children: [dataCell("cpc", 2600, GRAY_BG), dataCell("float ($)", 2000, GRAY_BG), dataCell("Cost Per Click = spend / clicks", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("cpm", 2600), dataCell("float ($)", 2000), dataCell("Cost Per 1000 Impressions = spend / impressions × 1000", 4760)] }),
      new TableRow({ children: [dataCell("cpp", 2600, GRAY_BG), dataCell("float ($)", 2000, GRAY_BG), dataCell("Cost Per 1000 People Reached", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("inline_link_clicks", 2600), dataCell("int", 2000), dataCell("نقرات الروابط المضمنة فقط (Destination Clicks)", 4760)] }),
      new TableRow({ children: [dataCell("inline_link_click_ctr", 2600, GRAY_BG), dataCell("percent", 2000, GRAY_BG), dataCell("CTR للروابط المضمنة فقط", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("inline_post_engagement", 2600), dataCell("int", 2000), dataCell("إجمالي التفاعل مع المنشور (لايك، كومنت، شير، إلخ)", 4760)] }),
      new TableRow({ children: [dataCell("social_spend", 2600, GRAY_BG), dataCell("float ($)", 2000, GRAY_BG), dataCell("الإنفاق على الإعلانات الاجتماعية (تظهر مع أسماء الأصدقاء)", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("buying_type", 2600), dataCell("string", 2000), dataCell("AUCTION أو RESERVED", 4760)] }),
      new TableRow({ children: [dataCell("objective", 2600, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("هدف الحملة", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("optimization_goal", 2600), dataCell("string", 2000), dataCell("هدف الأوبتيمايزيشن للأد سيت", 4760)] }),
      new TableRow({ children: [dataCell("date_start", 2600, GRAY_BG), dataCell("date", 2000, GRAY_BG), dataCell("تاريخ بداية الفترة المُبلَّغ عنها", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("date_stop", 2600), dataCell("date", 2000), dataCell("تاريخ نهاية الفترة المُبلَّغ عنها", 4760)] }),
    ]
  }),

  h3("2.6.2 مؤشرات التحويلات والأفعال"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2600, 2000, 4760],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 2600), hdrCell("النوع", 2000), hdrCell("الوصف", 4760)] }),
      new TableRow({ children: [dataCell("actions", 2600, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{action_type, value}] — جميع الأفعال مصنفة حسب النوع: purchase, lead, complete_registration, add_to_cart, initiate_checkout, view_content, search, add_payment_info, subscribe, contact, customize_product, donate, find_location, schedule, start_trial...", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("conversions", 2600), dataCell("array", 2000), dataCell("[{action_type, value}] — التحويلات المسجلة", 4760)] }),
      new TableRow({ children: [dataCell("cost_per_action_type", 2600, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{action_type, value}] — التكلفة لكل نوع عمل", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("cost_per_unique_action_type", 2600), dataCell("array", 2000), dataCell("[{action_type, value}] — تكلفة العمل الفريد", 4760)] }),
      new TableRow({ children: [dataCell("unique_actions", 2600, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{action_type, value}] — الأفعال من أشخاص فريدين", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("action_values", 2600), dataCell("array", 2000), dataCell("[{action_type, value}] — قيمة الأفعال (ROAS data) مجموع قيمة المشتريات", 4760)] }),
      new TableRow({ children: [dataCell("purchase_roas", 2600, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{action_type, value}] — Return On Ad Spend للمشتريات", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("conversion_values", 2600), dataCell("array", 2000), dataCell("[{action_type, value}] — قيمة التحويلات الإجمالية", 4760)] }),
      new TableRow({ children: [dataCell("cost_per_conversion", 2600, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{action_type, value}] — CPA لكل نوع تحويل", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("outbound_clicks", 2600), dataCell("array", 2000), dataCell("[{action_type, value}] — النقرات الخارجة من المنصة", 4760)] }),
      new TableRow({ children: [dataCell("outbound_clicks_ctr", 2600, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("[{action_type, value}] — معدل النقرات الخارجية", 4760, GRAY_BG)] }),
      new TableRow({ children: [dataCell("website_ctr", 2600), dataCell("array", 2000), dataCell("[{action_type, value}] — CTR الموقع الإلكتروني", 4760)] }),
    ]
  }),

  h3("2.6.3 Breakdowns — أبعاد التقسيم المتاحة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 6360],
    rows: [
      new TableRow({ children: [hdrCell("البُعد (Breakdown)", 3000), hdrCell("القيم المتاحة", 6360)] }),
      new TableRow({ children: [dataCell("age", 3000, GRAY_BG), dataCell("13-17, 18-24, 25-34, 35-44, 45-54, 55-64, 65+", 6360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("gender", 3000), dataCell("male, female, unknown", 6360)] }),
      new TableRow({ children: [dataCell("country", 3000, GRAY_BG), dataCell("رمز الدولة ISO-2 لكل دولة ظهر فيها الإعلان", 6360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("region", 3000), dataCell("المنطقة/الولاية داخل الدولة", 6360)] }),
      new TableRow({ children: [dataCell("dma", 3000, GRAY_BG), dataCell("Designated Market Area (المناطق الأمريكية فقط)", 6360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("impression_device", 3000), dataCell("desktop, mobile_app, mobile_web, ipad, iphone, ipod, android_tablet, android_smartphone", 6360)] }),
      new TableRow({ children: [dataCell("publisher_platform", 3000, GRAY_BG), dataCell("facebook, instagram, audience_network, messenger", 6360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("platform_position", 3000), dataCell("feed, right_hand_column, instant_article, marketplace, video_feeds, story, search, instream_video, reels, explore, rewarded_video", 6360)] }),
      new TableRow({ children: [dataCell("device_platform", 3000, GRAY_BG), dataCell("mobile, desktop", 6360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("product_id", 3000), dataCell("تقسيم حسب منتج في إعلانات الكتالوج DPA", 6360)] }),
      new TableRow({ children: [dataCell("hourly_stats_aggregated_by_advertiser_time_zone", 3000, GRAY_BG), dataCell("تقسيم بالساعة حسب المنطقة الزمنية للمعلن", 6360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("hourly_stats_aggregated_by_audience_time_zone", 3000), dataCell("تقسيم بالساعة حسب المنطقة الزمنية للجمهور", 6360)] }),
    ]
  }),
];

// ============================================================
// SECTION 3 - PLANNED FEATURES
// ============================================================
const sec3 = [
  pageBreak(),
  sectionBanner("القسم الثالث: الخطة المستقبلية والميزات الكاملة"),
  h1("3. هيكل التطبيق الكامل"),
  h2("3.1 معمارية النظام"),
  body("يُبنى التطبيق وفق معمارية Monorepo متعددة الطبقات تضمن مشاركة الكود بين المنصات وسهولة الصيانة:"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2800, 6560],
    rows: [
      new TableRow({ children: [hdrCell("الطبقة", 2800), hdrCell("التقنيات والمكونات", 6560)] }),
      new TableRow({ children: [dataCell("Core Library", 2800, GRAY_BG, true), dataCell("TypeScript SDK للـ Facebook Marketing API — مشترك بين جميع المنصات", 6560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("State Management", 2800, WHITE, true), dataCell("Zustand + React Query لإدارة الحالة والكاش والمزامنة", 6560)] }),
      new TableRow({ children: [dataCell("Web App", 2800, GRAY_BG, true), dataCell("Next.js 14 (App Router) + TailwindCSS + shadcn/ui", 6560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Desktop App", 2800, WHITE, true), dataCell("Electron أو Tauri — يستخدم نفس كود الـ Web", 6560)] }),
      new TableRow({ children: [dataCell("Mobile App", 2800, GRAY_BG, true), dataCell("React Native + Expo — iOS & Android من كودبيس واحد", 6560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("CLI Tool", 2800, WHITE, true), dataCell("Node.js CLI + Commander.js — للأتمتة وسكريبتات الـ Bulk", 6560)] }),
      new TableRow({ children: [dataCell("MCP Server", 2800, GRAY_BG, true), dataCell("Model Context Protocol Server للتكامل مع Claude AI", 6560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Backend API", 2800, WHITE, true), dataCell("FastAPI (Python) — للـ Webhooks والأتمتة والجدولة", 6560)] }),
      new TableRow({ children: [dataCell("Database", 2800, GRAY_BG, true), dataCell("PostgreSQL للبيانات الدائمة + Redis لكاش الـ Insights", 6560, GRAY_BG)] }),
    ]
  }),

  h2("3.2 خارطة الطريق — Roadmap"),
  h3("المرحلة الأولى (Phase 1): الأساس — الشهور 1-3"),
  callout("الهدف: تطبيق ويب يعمل بالكامل مع كل بيانات الـ API بشكل قرائي (Read)", BLUE_BG, BLUE_DARK),
  nb("تثبيت البنية التحتية: Monorepo (Nx/Turborepo)، CI/CD، إعدادات TypeScript الصارمة"),
  nb("OAuth 2.0 كامل مع Facebook — تسجيل دخول، تحديث Token، دعم حسابات متعددة"),
  nb("صفحة Dashboard رئيسية: KPIs لجميع الحسابات في نظرة واحدة"),
  nb("جدول Campaigns مع فلترة متقدمة، بحث، فرز، تصدير CSV"),
  nb("جدول Ad Sets مع عرض كامل لـ targeting وكل الحقول"),
  nb("جدول Ads مع معاينة Creative داخل الواجهة"),
  nb("نظام Insights: فلتر زمني مرن + Breakdowns + مقارنة فترات"),
  nb("تقارير أساسية قابلة للتصدير (PDF, Excel, CSV)"),

  h3("المرحلة الثانية (Phase 2): التحكم الكامل — الشهور 4-6"),
  callout("الهدف: Write Access — إنشاء وتعديل وحذف على جميع المستويات", ORANGE_BG, ORANGE),
  nb("إنشاء Campaigns مع كل خيارات الـ Objective والميزانية"),
  nb("إنشاء Ad Sets مع Targeting Builder بصري (خريطة، فلاتر، Lookalike)"),
  nb("إنشاء Ads مع Creative Builder: صور، فيديو، Carousel، Collection"),
  nb("Bulk Actions: تغيير الحالة، الميزانية، bid لمئات العناصر دفعة واحدة"),
  nb("Duplicate: نسخ Campaign/AdSet/Ad مع تعديلات"),
  nb("A/B Testing Framework: هيكل اختبار منظم مع مقارنة آلية"),
  nb("Budget Calculator: أداة تخطيط الميزانية قبل الإطلاق"),
  nb("Ad Preview: معاينة حقيقية في جميع المواضع (Feed, Story, Reels, Search...)"),

  h3("المرحلة الثالثة (Phase 3): الأتمتة الذكية — الشهور 7-9"),
  callout("الهدف: Rules Engine + AI-Powered Insights + Automated Optimization", GREEN_BG, GREEN),
  nb("Rules Engine: قواعد مشروطة (if CTR < 1% AND spend > $50 → PAUSE)"),
  nb("Auto Bidding: تعديل Bids تلقائياً لتحقيق هدف CPA محدد"),
  nb("Budget Pacing Monitor: تنبيه عند الإفراط أو التقتير في الإنفاق"),
  nb("Anomaly Detection: اكتشاف تغييرات مفاجئة في الأداء"),
  nb("AI Recommendations: اقتراحات تحسين مدعومة بـ Claude AI"),
  nb("Automated Reporting: إرسال تقارير يومية/أسبوعية تلقائياً بالبريد أو Slack"),
  nb("Audience Insights Integration: تقاطع بيانات Custom Audiences مع الأداء"),
  nb("Creative Performance Analysis: تحليل أداء الصور والفيديوهات"),

  h3("المرحلة الرابعة (Phase 4): التوسع — الشهور 10-12"),
  callout("الهدف: Multi-Platform Support + Agency Features + Advanced Analytics", ORANGE_BG, ORANGE),
  nb("إضافة Google Ads API: Campaigns, Ad Groups, Ads, Keywords"),
  nb("إضافة TikTok Ads API: Campaigns, Ad Groups, Ads"),
  nb("إضافة Snapchat Ads API"),
  nb("Unified Dashboard: مقارنة الأداء عبر جميع المنصات في مكان واحد"),
  nb("Client Portal: واجهة لعملاء الوكالات لرؤية تقاريرهم فقط"),
  nb("White-Label Mode: تغيير اللوجو والألوان لكل وكالة"),
  nb("Attribution Modeling: نماذج إسناد متعددة (First Touch, Last Touch, Linear, Data-Driven)"),
  nb("LTV & ROAS Forecasting: توقع قيمة العميل وعائد الإنفاق"),

  h2("3.3 الصفحات والشاشات الكاملة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2800, 2000, 4560],
    rows: [
      new TableRow({ children: [hdrCell("الصفحة", 2800), hdrCell("المرحلة", 2000), hdrCell("المحتوى والوظائف", 4560)] }),
      new TableRow({ children: [dataCell("/login", 2800, GRAY_BG), dataCell("Phase 1", 2000, GRAY_BG), dataCell("OAuth Facebook، حفظ Token، اختيار Business Account", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/dashboard", 2800), dataCell("Phase 1", 2000), dataCell("KPIs إجمالية، رسوم بيانية للأداء، تنبيهات، أنشطة حديثة", 4560)] }),
      new TableRow({ children: [dataCell("/accounts", 2800, GRAY_BG), dataCell("Phase 1", 2000, GRAY_BG), dataCell("كل الحسابات، إضافة حساب جديد، تفاصيل الفوترة", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/campaigns", 2800), dataCell("Phase 1", 2000), dataCell("جدول كامل، فلترة، بحث، Bulk Select", 4560)] }),
      new TableRow({ children: [dataCell("/campaigns/new", 2800, GRAY_BG), dataCell("Phase 2", 2000, GRAY_BG), dataCell("معالج إنشاء حملة خطوة بخطوة", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/campaigns/:id", 2800), dataCell("Phase 1", 2000), dataCell("تفاصيل حملة، مقاييسها، أد ستسها، خيارات تعديل", 4560)] }),
      new TableRow({ children: [dataCell("/adsets", 2800, GRAY_BG), dataCell("Phase 1", 2000, GRAY_BG), dataCell("جميع الأد ستس عبر الحسابات أو فلترة بالحملة", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/adsets/new", 2800), dataCell("Phase 2", 2000), dataCell("Targeting Builder متكامل مع خريطة جغرافية وفلاتر", 4560)] }),
      new TableRow({ children: [dataCell("/adsets/:id", 2800, GRAY_BG), dataCell("Phase 1", 2000, GRAY_BG), dataCell("تفاصيل أد سيت، إعلاناتها، أداؤها، تعديل الاستهداف", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/ads", 2800), dataCell("Phase 1", 2000), dataCell("جميع الإعلانات، فلتر بالحالة، معاينة مصغرة", 4560)] }),
      new TableRow({ children: [dataCell("/ads/new", 2800, GRAY_BG), dataCell("Phase 2", 2000, GRAY_BG), dataCell("Creative Builder: رفع صور/فيديو، Dynamic Creative، كتابة نص", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/ads/:id", 2800), dataCell("Phase 1", 2000), dataCell("تفاصيل الإعلان، المعاينة الحقيقية، أداؤه، issues", 4560)] }),
      new TableRow({ children: [dataCell("/analytics", 2800, GRAY_BG), dataCell("Phase 1", 2000, GRAY_BG), dataCell("Insights متقدمة، Breakdowns، مقارنة فترات، رسوم بيانية", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/audiences", 2800), dataCell("Phase 2", 2000), dataCell("Custom Audiences، Lookalike، إنشاء جمهور جديد، حجم الجمهور", 4560)] }),
      new TableRow({ children: [dataCell("/creatives", 2800, GRAY_BG), dataCell("Phase 2", 2000, GRAY_BG), dataCell("مكتبة الـ Creatives، أداء كل Creative، استخدامه في أد ستس", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/pixels", 2800), dataCell("Phase 2", 2000), dataCell("إدارة Facebook Pixels، اختبار الأحداث، إعداد الكونفيرجنس API", 4560)] }),
      new TableRow({ children: [dataCell("/rules", 2800, GRAY_BG), dataCell("Phase 3", 2000, GRAY_BG), dataCell("Rules Engine: إنشاء قواعد شرطية، تفعيل/تعطيل، سجل التنفيذ", 4560, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/reports", 2800), dataCell("Phase 3", 2000), dataCell("تقارير جاهزة + مخصصة، جدولة الإرسال، تصدير", 4560)] }),
      new TableRow({ children: [dataCell("/settings", 2800, GRAY_BG), dataCell("Phase 1", 2000, GRAY_BG), dataCell("إعدادات الحساب، التنبيهات، الفريق، الاشتراك", 4560, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// SECTION 4 - TECHNICAL SPECIFICATIONS
// ============================================================
const sec4 = [
  pageBreak(),
  sectionBanner("القسم الرابع: المواصفات التقنية"),
  h1("4. التكاملات والـ API Endpoints"),
  h2("4.1 Facebook Graph API — Endpoints الكاملة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3200, 2000, 4160],
    rows: [
      new TableRow({ children: [hdrCell("الـ Endpoint", 3200), hdrCell("Method", 2000), hdrCell("الاستخدام", 4160)] }),
      new TableRow({ children: [dataCell("/me/adaccounts", 3200, GRAY_BG), dataCell("GET", 2000, GRAY_BG), dataCell("جلب جميع حسابات الإعلان المرتبطة", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/act_{id}", 3200), dataCell("GET / POST", 2000), dataCell("تفاصيل حساب + تعديل الإعدادات", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/campaigns", 3200, GRAY_BG), dataCell("GET / POST", 2000, GRAY_BG), dataCell("جلب الحملات / إنشاء حملة جديدة", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/{campaign_id}", 3200), dataCell("GET / POST / DELETE", 2000), dataCell("تفاصيل / تعديل / حذف حملة", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/adsets", 3200, GRAY_BG), dataCell("GET / POST", 2000, GRAY_BG), dataCell("جلب الأد ستس / إنشاء أد سيت", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/{adset_id}", 3200), dataCell("GET / POST / DELETE", 2000), dataCell("تفاصيل / تعديل / حذف أد سيت", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/ads", 3200, GRAY_BG), dataCell("GET / POST", 2000, GRAY_BG), dataCell("جلب الإعلانات / إنشاء إعلان", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/{ad_id}", 3200), dataCell("GET / POST / DELETE", 2000), dataCell("تفاصيل / تعديل / حذف إعلان", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/adcreatives", 3200, GRAY_BG), dataCell("GET / POST", 2000, GRAY_BG), dataCell("جلب / إنشاء Creatives", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/{creative_id}", 3200), dataCell("GET / DELETE", 2000), dataCell("تفاصيل / حذف Creative", 4160)] }),
      new TableRow({ children: [dataCell("/{campaign_id}/insights", 3200, GRAY_BG), dataCell("GET / POST (async)", 2000, GRAY_BG), dataCell("Insights للحملة (sync أو async report)", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/{adset_id}/insights", 3200), dataCell("GET / POST (async)", 2000), dataCell("Insights للأد سيت", 4160)] }),
      new TableRow({ children: [dataCell("/{ad_id}/insights", 3200, GRAY_BG), dataCell("GET / POST (async)", 2000, GRAY_BG), dataCell("Insights للإعلان الفردي", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/act_{id}/insights", 3200), dataCell("GET / POST (async)", 2000), dataCell("Insights على مستوى الحساب كله", 4160)] }),
      new TableRow({ children: [dataCell("/{report_run_id}", 3200, GRAY_BG), dataCell("GET", 2000, GRAY_BG), dataCell("التحقق من حالة تقرير Async ونتائجه", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/act_{id}/customaudiences", 3200), dataCell("GET / POST", 2000), dataCell("إدارة Custom Audiences", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/adimages", 3200, GRAY_BG), dataCell("GET / POST", 2000, GRAY_BG), dataCell("رفع وإدارة صور الإعلانات", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/act_{id}/advideos", 3200), dataCell("GET / POST", 2000), dataCell("رفع وإدارة فيديوهات الإعلانات", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/activities", 3200, GRAY_BG), dataCell("GET", 2000, GRAY_BG), dataCell("سجل التغييرات على الحساب", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/act_{id}/reachestimate", 3200), dataCell("GET", 2000), dataCell("تقدير الوصول لإعدادات استهداف محددة", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/delivery_estimate", 3200, GRAY_BG), dataCell("GET", 2000, GRAY_BG), dataCell("تقدير التوصيل للأد سيت قبل الإطلاق", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("/search?type=adinterest", 3200), dataCell("GET", 2000), dataCell("البحث عن اهتمامات الاستهداف", 4160)] }),
      new TableRow({ children: [dataCell("/act_{id}/adsets (batch)", 3200, GRAY_BG), dataCell("POST (Batch API)", 2000, GRAY_BG), dataCell("Batch Request: تنفيذ 50 عملية في طلب واحد", 4160, GRAY_BG)] }),
    ]
  }),

  h2("4.2 حالات الأمان والصلاحيات"),
  h3("Facebook Permissions المطلوبة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("الصلاحية", 3000), hdrCell("النوع", 2000), hdrCell("الاستخدام", 4360)] }),
      new TableRow({ children: [dataCell("ads_read", 3000, GRAY_BG), dataCell("Read", 2000, GRAY_BG), dataCell("قراءة بيانات الحملات والأد ستس والإعلانات والـ Insights", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("ads_management", 3000), dataCell("Read / Write", 2000), dataCell("إنشاء وتعديل وحذف الحملات والأد ستس والإعلانات", 4360)] }),
      new TableRow({ children: [dataCell("business_management", 3000, GRAY_BG), dataCell("Read / Write", 2000, GRAY_BG), dataCell("إدارة Business Manager والوصول لحسابات العملاء", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("pages_read_engagement", 3000), dataCell("Read", 2000), dataCell("قراءة بيانات الصفحات المرتبطة بالإعلانات", 4360)] }),
      new TableRow({ children: [dataCell("instagram_basic", 3000, GRAY_BG), dataCell("Read", 2000, GRAY_BG), dataCell("بيانات حسابات انستغرام المرتبطة", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("leads_retrieval", 3000), dataCell("Read", 2000), dataCell("استرجاع بيانات العملاء المحتملين من Lead Ads", 4360)] }),
    ]
  }),

  h2("4.3 المنصات المستهدفة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2000, 2000, 2680, 2680],
    rows: [
      new TableRow({ children: [hdrCell("المنصة", 2000), hdrCell("Framework", 2000), hdrCell("المميزات الإضافية", 2680), hdrCell("المرحلة", 2680)] }),
      new TableRow({ children: [dataCell("Web App", 2000, GRAY_BG), dataCell("Next.js 14", 2000, GRAY_BG), dataCell("SSR, PWA, Keyboard Shortcuts, Dark Mode", 2680, GRAY_BG), dataCell("Phase 1", 2680, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Desktop (Win/Mac/Linux)", 2000), dataCell("Tauri v2", 2000), dataCell("Native Notifications, Offline Mode, System Tray", 2680), dataCell("Phase 2", 2680)] }),
      new TableRow({ children: [dataCell("iOS & Android", 2000, GRAY_BG), dataCell("React Native + Expo", 2000, GRAY_BG), dataCell("Push Notifications, Biometric Auth, Widget", 2680, GRAY_BG), dataCell("Phase 3", 2680, GRAY_BG)] }),
      new TableRow({ children: [dataCell("CLI", 2000), dataCell("Node.js", 2000), dataCell("Bulk Operations, Scripting, Cron Jobs", 2680), dataCell("Phase 2", 2680)] }),
      new TableRow({ children: [dataCell("Browser Extension", 2000, GRAY_BG), dataCell("WXT Framework", 2000, GRAY_BG), dataCell("Quick Stats على أي صفحة، إضافة سريعة للـ Audiences", 2680, GRAY_BG), dataCell("Phase 4", 2680, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Slack Bot", 2000), dataCell("Bolt.js", 2000), dataCell("تقارير يومية، تنبيهات، أوامر /fbads", 2680), dataCell("Phase 3", 2680)] }),
      new TableRow({ children: [dataCell("Claude MCP", 2000, GRAY_BG), dataCell("MCP Server", 2000, GRAY_BG), dataCell("التحكم الكامل بالأوامر الطبيعية عبر Claude AI", 2680, GRAY_BG), dataCell("Phase 1", 2680, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// SECTION 5 - UX/UI DESIGN
// ============================================================
const sec5 = [
  pageBreak(),
  sectionBanner("القسم الخامس: تصميم تجربة المستخدم"),
  h1("5. تجربة المستخدم والتصميم"),
  h2("5.1 مبادئ التصميم"),
  bullet("Data First: البيانات دائماً في المقدمة — لا خطوات إضافية للوصول إليها"),
  bullet("Progressive Disclosure: إظهار التفاصيل تدريجياً — البسيط أولاً ثم المتقدم"),
  bullet("Bulk by Default: كل قائمة تدعم تحديد متعدد وعمليات جماعية"),
  bullet("Real-time Feedback: كل تغيير يظهر نتيجته فوراً"),
  bullet("Keyboard-First: اختصارات لوحة مفاتيح لكل عملية رئيسية"),
  bullet("Multilingual: دعم كامل للعربية والإنجليزية مع RTL"),
  bullet("Accessible: WCAG 2.1 AA — دعم قراء الشاشة، تباين الألوان"),

  h2("5.2 نظام الألوان والـ Design System"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2400, 2000, 4960],
    rows: [
      new TableRow({ children: [hdrCell("الرمز", 2400), hdrCell("اللون", 2000), hdrCell("الاستخدام", 4960)] }),
      new TableRow({ children: [dataCell("primary", 2400, GRAY_BG), dataCell("#1877F2 (FB Blue)", 2000, GRAY_BG), dataCell("الأزرار الرئيسية، الروابط، النشط", 4960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("success", 2400), dataCell("#42B72A", 2000), dataCell("الحالة النشطة، التأكيدات، الأرقام الإيجابية", 4960)] }),
      new TableRow({ children: [dataCell("warning", 2400, GRAY_BG), dataCell("#F5A623", 2000, GRAY_BG), dataCell("التحذيرات، ضعف الأداء، الميزانية المنخفضة", 4960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("danger", 2400), dataCell("#FA3E3E", 2000), dataCell("الأخطاء، الإيقاف، الحذف، المشاكل", 4960)] }),
      new TableRow({ children: [dataCell("neutral", 2400, GRAY_BG), dataCell("#65676B", 2000, GRAY_BG), dataCell("النصوص الثانوية، الحدود، الخلفيات", 4960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("dark", 2400), dataCell("#1C1E21", 2000), dataCell("النصوص الرئيسية، الرؤوس", 4960)] }),
    ]
  }),

  h2("5.3 مكونات واجهة المستخدم الرئيسية"),
  h4("Quick Stats Bar"),
  body("شريط ثابت في أعلى الصفحة يظهر: إجمالي الإنفاق اليوم، عدد الحملات النشطة، CTR الإجمالي، تنبيهات طارئة. يتحدث كل 60 ثانية."),

  h4("Campaign Tree View"),
  body("عرض شجري هرمي: Account → Campaign → Ad Set → Ad. كل مستوى قابل للطي/التوسيع. الضغط على أي عنصر يفتح لوحة التفاصيل الجانبية بدون مغادرة الصفحة."),

  h4("Bulk Actions Toolbar"),
  body("يظهر عند تحديد أكثر من عنصر. يتضمن: تشغيل/إيقاف/أرشفة، تغيير الميزانية بنسبة مئوية أو قيمة ثابتة، نسخ، نقل لحملة أخرى، تصدير."),

  h4("Insights Chart"),
  body("رسم بياني تفاعلي متعدد المحاور. المحور Y الأيسر للإنفاق، الأيمن لـ CTR. دعم تحديد نطاق زمني بالسحب. مقارنة فترتين بألوان مختلفة. Export كـ PNG أو CSV."),

  h4("Targeting Visualizer"),
  body("عرض بصري لإعدادات الاستهداف: خريطة جغرافية للمناطق المستهدفة، Venn Diagram للتقاطعات والاستثناءات، حجم الجمهور المقدر في الوقت الحقيقي."),
];

// ============================================================
// SECTION 6 - CURRENT MCP STATE
// ============================================================
const sec6 = [
  pageBreak(),
  sectionBanner("القسم السادس: الحالة الراهنة والتوثيق"),
  h1("6. الـ MCP Server الحالي"),
  h2("6.1 الأدوات المتاحة حالياً"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3500, 5860],
    rows: [
      new TableRow({ children: [hdrCell("اسم الأداة (Tool)", 3500), hdrCell("الوصف", 5860)] }),
      new TableRow({ children: [dataCell("list_ad_accounts", 3500, GRAY_BG), dataCell("جلب جميع حسابات الإعلان المرتبطة بالـ Token", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_details_of_ad_account", 3500), dataCell("تفاصيل حساب إعلاني محدد (status, spent, currency, balance)", 5860)] }),
      new TableRow({ children: [dataCell("get_adaccount_insights", 3500, GRAY_BG), dataCell("Insights كاملة على مستوى الحساب مع دعم Breakdowns والفترات الزمنية", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_campaigns_by_adaccount", 3500), dataCell("جلب جميع الحملات في حساب مع فلترة بالحالة والهدف", 5860)] }),
      new TableRow({ children: [dataCell("get_campaign_by_id", 3500, GRAY_BG), dataCell("تفاصيل كاملة لحملة محددة بجميع حقولها", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_campaign_insights", 3500), dataCell("Insights لحملة بعينها مع Breakdowns ومقارنة فترات", 5860)] }),
      new TableRow({ children: [dataCell("get_adsets_by_adaccount", 3500, GRAY_BG), dataCell("جلب جميع الأد ستس في حساب", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_adsets_by_campaign", 3500), dataCell("جلب الأد ستس التابعة لحملة محددة", 5860)] }),
      new TableRow({ children: [dataCell("get_adset_by_id", 3500, GRAY_BG), dataCell("تفاصيل كاملة لأد سيت (targeting, budget, schedule, optimization)", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_adset_insights", 3500), dataCell("Insights لأد سيت محدد", 5860)] }),
      new TableRow({ children: [dataCell("get_adsets_by_ids", 3500, GRAY_BG), dataCell("جلب عدة أد ستس دفعة واحدة بقائمة IDs", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_ads_by_adaccount", 3500), dataCell("جلب جميع الإعلانات في حساب مع فلتر الحالة", 5860)] }),
      new TableRow({ children: [dataCell("get_ads_by_campaign", 3500, GRAY_BG), dataCell("جلب الإعلانات التابعة لحملة", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_ads_by_adset", 3500), dataCell("جلب الإعلانات التابعة لأد سيت", 5860)] }),
      new TableRow({ children: [dataCell("get_ad_by_id", 3500, GRAY_BG), dataCell("تفاصيل إعلان محدد (creative, status, issues, preview_link)", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_ad_creative_by_id", 3500), dataCell("تفاصيل الـ Creative (نص، صورة، رابط، CTA)", 5860)] }),
      new TableRow({ children: [dataCell("get_ad_creatives_by_ad_id", 3500, GRAY_BG), dataCell("جلب Creatives مرتبطة بإعلان", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_ad_insights", 3500), dataCell("Insights لإعلان فردي", 5860)] }),
      new TableRow({ children: [dataCell("get_activities_by_adaccount", 3500, GRAY_BG), dataCell("سجل التغييرات الأخيرة على الحساب", 5860, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_activities_by_adset", 3500), dataCell("سجل التغييرات على أد سيت محدد", 5860)] }),
      new TableRow({ children: [dataCell("fetch_pagination_url", 3500, GRAY_BG), dataCell("استرجاع صفحة تالية من Cursor-Based Pagination", 5860, GRAY_BG)] }),
    ]
  }),

  h2("6.2 الأدوات المخططة للإضافة"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3200, 2000, 4160],
    rows: [
      new TableRow({ children: [hdrCell("الأداة المطلوبة", 3200), hdrCell("الأولوية", 2000), hdrCell("الاستخدام", 4160)] }),
      new TableRow({ children: [dataCell("create_campaign", 3200, GRAY_BG), dataCell("عالية جداً", 2000, RED_BG, false, RED), dataCell("إنشاء حملة جديدة بجميع الإعدادات", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("update_campaign", 3200), dataCell("عالية جداً", 2000, RED_BG, false, RED), dataCell("تعديل حملة (اسم، ميزانية، حالة، bid strategy)", 4160)] }),
      new TableRow({ children: [dataCell("create_adset", 3200, GRAY_BG), dataCell("عالية جداً", 2000, RED_BG, false, RED), dataCell("إنشاء أد سيت جديد مع كل إعدادات الاستهداف", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("update_adset", 3200), dataCell("عالية جداً", 2000, RED_BG, false, RED), dataCell("تعديل أد سيت (استهداف، ميزانية، جدولة، حالة)", 4160)] }),
      new TableRow({ children: [dataCell("create_ad", 3200, GRAY_BG), dataCell("عالية جداً", 2000, RED_BG, false, RED), dataCell("إنشاء إعلان جديد مع Creative ومعلمات تتبع", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("update_ad_status", 3200), dataCell("عالية جداً", 2000, RED_BG, false, RED), dataCell("تغيير حالة إعلان (ACTIVE/PAUSED/DELETED)", 4160)] }),
      new TableRow({ children: [dataCell("create_ad_creative", 3200, GRAY_BG), dataCell("عالية", 2000, ORANGE_BG, false, ORANGE), dataCell("إنشاء Creative جديد (صورة/فيديو/Carousel)", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("bulk_update_status", 3200), dataCell("عالية", 2000, ORANGE_BG, false, ORANGE), dataCell("Batch API لتغيير حالة عشرات العناصر دفعة واحدة", 4160)] }),
      new TableRow({ children: [dataCell("duplicate_campaign", 3200, GRAY_BG), dataCell("عالية", 2000, ORANGE_BG, false, ORANGE), dataCell("نسخ حملة كاملة مع كل أد ستسها وإعلاناتها", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_audiences", 3200), dataCell("متوسطة", 2000, GREEN_BG, false, GREEN), dataCell("جلب Custom Audiences و Lookalike Audiences", 4160)] }),
      new TableRow({ children: [dataCell("create_custom_audience", 3200, GRAY_BG), dataCell("متوسطة", 2000, GREEN_BG, false, GREEN), dataCell("إنشاء Custom Audience من Customer List أو Pixel", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_reach_estimate", 3200), dataCell("متوسطة", 2000, GREEN_BG, false, GREEN), dataCell("تقدير حجم الجمهور لإعدادات استهداف معينة", 4160)] }),
      new TableRow({ children: [dataCell("upload_ad_image", 3200, GRAY_BG), dataCell("متوسطة", 2000, GREEN_BG, false, GREEN), dataCell("رفع صورة إلى مكتبة صور الحساب", 4160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("get_ad_account_billing", 3200), dataCell("منخفضة", 2000, GRAY_BG), dataCell("تفاصيل الفواتير وتاريخ الدفعات", 4160)] }),
      new TableRow({ children: [dataCell("create_automated_rule", 3200, GRAY_BG), dataCell("منخفضة", 2000, GRAY_BG), dataCell("إنشاء قواعد تلقائية على مستوى Facebook", 4160, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// SECTION 7 - GLOSSARY & APPENDIX
// ============================================================
const sec7 = [
  pageBreak(),
  sectionBanner("القسم السابع: المسرد والملاحق"),
  h1("7. مسرد المصطلحات"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2200, 7160],
    rows: [
      new TableRow({ children: [hdrCell("المصطلح", 2200), hdrCell("التعريف", 7160)] }),
      new TableRow({ children: [dataCell("Ad Account", 2200, GRAY_BG, true), dataCell("الحساب الإعلاني — الحاوية الرئيسية التي تحتوي على جميع الحملات وتربط بمصدر الدفع", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Campaign", 2200, WHITE, true), dataCell("الحملة الإعلانية — تحدد الهدف التسويقي (Objective) والميزانية الكلية واستراتيجية المزايدة", 7160)] }),
      new TableRow({ children: [dataCell("Ad Set", 2200, GRAY_BG, true), dataCell("المجموعة الإعلانية — تحدد الجمهور المستهدف والجدولة والميزانية اليومية ومنصات العرض", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Ad", 2200, WHITE, true), dataCell("الإعلان الفردي — يحتوي على الـ Creative (الصورة/الفيديو/النص) ورابط الوجهة وزر CTA", 7160)] }),
      new TableRow({ children: [dataCell("Creative", 2200, GRAY_BG, true), dataCell("المحتوى الإبداعي — الصورة أو الفيديو أو النص أو الكاروسيل الذي يراه المستخدم", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Objective", 2200, WHITE, true), dataCell("الهدف التسويقي للحملة — ماذا تريد أن يفعل المستخدم (شراء، نقرة، مشاهدة، تسجيل...)", 7160)] }),
      new TableRow({ children: [dataCell("Optimization Goal", 2200, GRAY_BG, true), dataCell("هدف الأوبتيمايزيشن — ما يحاول Facebook تعظيمه عند توزيع الميزانية (نقرات، تحويلات، مشاهدات...)", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Bid Strategy", 2200, WHITE, true), dataCell("استراتيجية المزايدة — كيف يقدم Facebook عطاءات في المزادات (الأرخص، بحد أقصى، بهدف تكلفة)", 7160)] }),
      new TableRow({ children: [dataCell("CPC", 2200, GRAY_BG, true), dataCell("Cost Per Click — التكلفة لكل نقرة = الإنفاق ÷ عدد النقرات", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("CPM", 2200, WHITE, true), dataCell("Cost Per Mille — التكلفة لكل 1000 مشاهدة = الإنفاق ÷ المشاهدات × 1000", 7160)] }),
      new TableRow({ children: [dataCell("CTR", 2200, GRAY_BG, true), dataCell("Click-Through Rate — معدل النقر = النقرات ÷ المشاهدات × 100%", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("ROAS", 2200, WHITE, true), dataCell("Return On Ad Spend — عائد الإنفاق الإعلاني = إجمالي قيمة المبيعات ÷ الإنفاق الإعلاني", 7160)] }),
      new TableRow({ children: [dataCell("CPA", 2200, GRAY_BG, true), dataCell("Cost Per Action — التكلفة لكل عمل = الإنفاق ÷ عدد التحويلات", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("CPP", 2200, WHITE, true), dataCell("Cost Per Person — التكلفة للوصول لكل 1000 شخص فريد", 7160)] }),
      new TableRow({ children: [dataCell("Frequency", 2200, GRAY_BG, true), dataCell("التكرار — متوسط مرات رؤية الإعلان لكل شخص في الفترة المحددة", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Reach", 2200, WHITE, true), dataCell("الوصول — عدد الأشخاص الفريدين الذين رأوا الإعلان مرة واحدة على الأقل", 7160)] }),
      new TableRow({ children: [dataCell("Impression", 2200, GRAY_BG, true), dataCell("المشاهدة — كل مرة يظهر فيها الإعلان على شاشة مستخدم (يختلف عن Reach)", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Breakdown", 2200, WHITE, true), dataCell("التقسيم — تجزئة بيانات الأداء حسب عمر أو جنس أو دولة أو جهاز أو موضع إعلاني", 7160)] }),
      new TableRow({ children: [dataCell("Lookalike Audience", 2200, GRAY_BG, true), dataCell("جمهور مشابه — يجده Facebook بمواصفات مشابهة لقائمة عملائك الحاليين (1%-10% من سكان دولة)", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Pixel", 2200, WHITE, true), dataCell("كود JavaScript يُنصب على الموقع لتتبع أحداث المستخدمين (زيارة، شراء، سلة...) وربطها بالإعلانات", 7160)] }),
      new TableRow({ children: [dataCell("Attribution Window", 2200, GRAY_BG, true), dataCell("نافذة الإسناد — الفترة الزمنية بعد النقرة أو المشاهدة لاحتساب التحويل كنتيجة للإعلان", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("DPA", 2200, WHITE, true), dataCell("Dynamic Product Ads — إعلانات تعرض منتجات من كتالوج بشكل تلقائي بناءً على سلوك المستخدم", 7160)] }),
      new TableRow({ children: [dataCell("CBO", 2200, GRAY_BG, true), dataCell("Campaign Budget Optimization — يوزع Facebook الميزانية تلقائياً بين الأد ستس حسب الأداء", 7160, GRAY_BG)] }),
      new TableRow({ children: [dataCell("SKAdNetwork", 2200, WHITE, true), dataCell("إطار Apple لتتبع تثبيتات التطبيقات على iOS 14.5+ مع احترام الخصوصية", 7160)] }),
      new TableRow({ children: [dataCell("Conversion API", 2200, GRAY_BG, true), dataCell("إرسال أحداث التحويل مباشرة من السيرفر إلى Facebook بدون الاعتماد على Pixel في المتصفح", 7160, GRAY_BG)] }),
    ]
  }),

  h1("8. ملاحق"),
  h2("8.1 قائمة مرجعية لأهداف الحملات"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [2400, 3000, 3960],
    rows: [
      new TableRow({ children: [hdrCell("الـ Objective", 2400), hdrCell("مرحلة القمع (Funnel)", 3000), hdrCell("الاستخدام الأمثل", 3960)] }),
      new TableRow({ children: [dataCell("BRAND_AWARENESS", 2400, GRAY_BG), dataCell("Top of Funnel", 3000, GRAY_BG), dataCell("زيادة الوعي بعلامة تجارية جديدة في السوق", 3960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("REACH", 2400), dataCell("Top of Funnel", 3000), dataCell("الوصول لأكبر عدد من الأشخاص بتكلفة منخفضة", 3960)] }),
      new TableRow({ children: [dataCell("VIDEO_VIEWS", 2400, GRAY_BG), dataCell("Top of Funnel", 3000, GRAY_BG), dataCell("تعزيز محتوى الفيديو والحصول على مشاهدات", 3960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("POST_ENGAGEMENT", 2400), dataCell("Mid Funnel", 3000), dataCell("زيادة اللايكات والكومنتات والمشاركات على المنشورات", 3960)] }),
      new TableRow({ children: [dataCell("PAGE_LIKES", 2400, GRAY_BG), dataCell("Mid Funnel", 3000, GRAY_BG), dataCell("زيادة متابعي صفحة الفيسبوك", 3960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("LINK_CLICKS", 2400), dataCell("Mid Funnel", 3000), dataCell("توليد زيارات للموقع الإلكتروني", 3960)] }),
      new TableRow({ children: [dataCell("MESSAGES", 2400, GRAY_BG), dataCell("Mid Funnel", 3000, GRAY_BG), dataCell("بدء محادثات في Messenger أو WhatsApp أو Instagram", 3960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("EVENT_RESPONSES", 2400), dataCell("Mid Funnel", 3000), dataCell("الترويج لأحداث والحصول على ردود", 3960)] }),
      new TableRow({ children: [dataCell("APP_INSTALLS", 2400, GRAY_BG), dataCell("Bottom of Funnel", 3000, GRAY_BG), dataCell("زيادة تحميلات التطبيق على iOS وAndroid", 3960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("LEAD_GENERATION", 2400), dataCell("Bottom of Funnel", 3000), dataCell("جمع بيانات العملاء المحتملين عبر Lead Form مدمج", 3960)] }),
      new TableRow({ children: [dataCell("CONVERSIONS", 2400, GRAY_BG), dataCell("Bottom of Funnel", 3000, GRAY_BG), dataCell("تحفيز شراء أو تسجيل على الموقع (يحتاج Pixel)", 3960, GRAY_BG)] }),
      new TableRow({ children: [dataCell("PRODUCT_CATALOG_SALES", 2400), dataCell("Bottom of Funnel", 3000), dataCell("إعلانات الكتالوج الديناميكية DPA للتجارة الإلكترونية", 3960)] }),
      new TableRow({ children: [dataCell("STORE_VISITS", 2400, GRAY_BG), dataCell("Offline", 3000, GRAY_BG), dataCell("توليد زيارات للمتاجر الفعلية (يحتاج Store Locations)", 3960, GRAY_BG)] }),
    ]
  }),

  h2("8.2 سجل الإصدارات"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [1500, 1500, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("الإصدار", 1500), hdrCell("التاريخ", 1500), hdrCell("الكاتب", 2000), hdrCell("التغييرات", 4360)] }),
      new TableRow({ children: [dataCell("v1.0", 1500, GRAY_BG), dataCell("مايو 2026", 1500, GRAY_BG), dataCell("الفريق", 2000, GRAY_BG), dataCell("الإصدار الأول الشامل — توثيق كامل لـ Facebook API + Roadmap + MCP State", 4360, GRAY_BG)] }),
    ]
  }),

  ...empty(2),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 400, after: 80 },
    children: [new TextRun({ text: "نهاية الوثيقة", bold: true, size: 24, color: TEXT_MID, font: "Arial" })]
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Facebook Ads Multi-Platform Manager GDD v1.0 · مايو 2026", size: 18, color: TEXT_MID, font: "Arial" })]
  }),
];

// ============================================================
// SECTION 8 - BUSINESS MANAGER & ASSETS
// ============================================================
const sec8 = [
  pageBreak(),
  sectionBanner("القسم الثامن: Business Manager والأصول (Assets)"),
  h1("8. إدارة Business Manager (BM)"),
  h2("8.1 حسابات الأعمال (Business Accounts)"),
  body("الـ Business Manager هو المظلة التي تجمع الحسابات الإعلانية، الصفحات، البيكسلات، والأشخاص."),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 3000), hdrCell("النوع", 2000), hdrCell("الوصف", 4360)] }),
      new TableRow({ children: [dataCell("id", 3000, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("معرف مدير الأعمال", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 3000), dataCell("string", 2000), dataCell("اسم النشاط التجاري", 4360)] }),
      new TableRow({ children: [dataCell("verification_status", 3000, GRAY_BG), dataCell("enum", 2000, GRAY_BG), dataCell("حالة توثيق النشاط (verified, unverified)", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("creation_time", 3000), dataCell("datetime", 2000), dataCell("تاريخ إنشاء الحساب", 4360)] }),
    ]
  }),
  h2("8.2 الصفحات وحسابات انستغرام وواتساب"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("المورد", 3000), hdrCell("API Endpoint", 2000), hdrCell("الاستخدام المتاح", 4360)] }),
      new TableRow({ children: [dataCell("Facebook Pages", 3000, GRAY_BG), dataCell("/{business_id}/owned_pages", 2000, GRAY_BG), dataCell("جلب صفحات فيسبوك، نشر محتوى، استخدامها في الإعلانات", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Instagram Accounts", 3000), dataCell("/{business_id}/instagram_accounts", 2000), dataCell("جلب حسابات انستغرام وربطها بالصفحات للإعلانات", 4360)] }),
      new TableRow({ children: [dataCell("WhatsApp Business", 3000, GRAY_BG), dataCell("/{business_id}/whatsapp_business_accounts", 2000, GRAY_BG), dataCell("جلب أرقام واتساب للأعمال لاستخدامها في إعلانات Click-to-WhatsApp", 4360, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// SECTION 9 - ADVANCED TRACKING & CATALOGS
// ============================================================
const sec9 = [
  pageBreak(),
  sectionBanner("القسم التاسع: التتبع المتقدم والكتالوجات"),
  h1("9. Pixels & Conversions API (CAPI)"),
  h2("9.1 Facebook Pixels"),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("الخاصية / الحقل", 3000), hdrCell("النوع", 2000), hdrCell("الوصف", 4360)] }),
      new TableRow({ children: [dataCell("id", 3000, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("معرف البيكسل", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("name", 3000), dataCell("string", 2000), dataCell("اسم البيكسل", 4360)] }),
      new TableRow({ children: [dataCell("rules", 3000, GRAY_BG), dataCell("object", 2000, GRAY_BG), dataCell("قواعد Custom Conversions", 4360, GRAY_BG)] }),
    ]
  }),
  body("سيدعم التطبيق مستقبلاً إرسال الأحداث مباشرة من السيرفر باستخدام Conversions API (CAPI) لضمان دقة التتبع مع تحديثات الخصوصية (iOS 14+)."),
  
  h1("10. Product Catalogs"),
  body("دعم كامل لإدارة الكتالوجات (Product Feeds) لاستخدامها في إعلانات DPA."),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("المورد", 3000), hdrCell("Endpoint", 2000), hdrCell("الوصف", 4360)] }),
      new TableRow({ children: [dataCell("Catalogs", 3000, GRAY_BG), dataCell("/{business_id}/product_catalogs", 2000, GRAY_BG), dataCell("إنشاء وجلب الكتالوجات", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("Product Sets", 3000), dataCell("/{catalog_id}/product_sets", 2000), dataCell("مجموعات المنتجات المفلترة (مثال: أحذية رياضية أقل من 50$)", 4360)] }),
      new TableRow({ children: [dataCell("Products", 3000, GRAY_BG), dataCell("/{catalog_id}/products", 2000, GRAY_BG), dataCell("المنتجات الفردية، أسعارها، صورها، وحالتها (متوفر/نفذ)", 4360, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// SECTION 10 - ULTIMATE CROSS-PLATFORM ROADMAP
// ============================================================
const sec10 = [
  pageBreak(),
  sectionBanner("القسم العاشر: خطة الهيمنة الشاملة للتطبيق"),
  h1("11. الخطط المستقبلية لكل المنصات"),
  h2("11.1 منصات التحكم (Control Planes)"),
  body("سيصبح التطبيق النظام المركزي الشامل لإدارة الإعلانات، متخطياً قيود واجهة فيسبوك التقليدية، من خلال توفير نقطة وصول تناسب كل حالة استخدام:"),
  bullet("1. الويب (Web Dashboard): الواجهة الرئيسية للمديرين والـ Media Buyers للتحليل العميق والـ Bulk Actions."),
  bullet("2. تطبيق الديسكتوب (Desktop App): يعمل في الخلفية (System Tray)، تنبيهات لحظية للصرف الزائد، أسرع في الأداء باستهلاك أقل للرام."),
  bullet("3. تطبيق الموبايل (Mobile App): للتحكم الطارئ (إيقاف حملة، زيادة ميزانية، التحقق من العائد ROAS) وأنت بعيد عن مكتبك."),
  bullet("4. إضافة المتصفح (Browser Extension): تظهر إحصائيات سريعة للبيكسل في موقعك، وتحليل إعلانات المنافسين (Ad Library Analyzer)."),
  bullet("5. أداة سطر الأوامر (CLI Tool): للمطورين لأتمتة إنشاء مئات الحملات باستخدام سكريبتات أو ملفات CSV."),
  bullet("6. الـ MCP Server (Claude Integration): التحكم بالحملات وتوليد تقارير وسؤال أسئلة عن الأداء باللغة الطبيعية!"),
  
  h2("11.2 الأتمتة المتقدمة (Automated Rules Engine)"),
  body("النظام القادم سيمتلك محرك قواعد متقدم يعمل على السيرفر الخاص بنا (وليس قواعد فيسبوك المحدودة):"),
  bullet("Stop-Loss Rules: إيقاف الإعلان إذا صرف 20$ بدون أي مبيعات."),
  bullet("Scale Rules: زيادة الميزانية 20% يومياً إذا كان الـ ROAS أعلى من 3.0."),
  bullet("Dayparting: تفعيل الإعلانات وإيقافها في ساعات محددة بناءً على الوقت الأكثر ربحية."),
  bullet("Weather-Based Ads: تفعيل إعلانات معينة بناءً على حالة الطقس في المنطقة المستهدفة (باستخدام Weather API)."),

  h2("11.3 الذكاء الاصطناعي (AI Capabilities)"),
  bullet("توليد الـ Creatives: استخدام AI لتوليد نصوص إعلانية وصور بديلة للحملات الخاسرة."),
  bullet("Predictive Analytics: توقع الأداء المستقبلي للحملة بناءً على بيانات أول 24 ساعة."),
  bullet("Audience Generator: اكتشاف اهتمامات خفية (Hidden Interests) غير مقترحة في واجهة فيسبوك."),
];

// ============================================================
// SECTION 11 - WEBHOOKS & REAL-TIME UPDATES
// ============================================================
const sec11 = [
  pageBreak(),
  sectionBanner("القسم الحادي عشر: التحديثات اللحظية (Webhooks)"),
  h1("12. Facebook Webhooks"),
  body("لتحقيق أصغر تحكم ممكن، سيعتمد التطبيق على Webhooks لاستقبال التحديثات فور حدوثها على فيسبوك دون الحاجة لعمل Polling مستمر."),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("الموضوع (Topic)", 3000), hdrCell("الحدث (Event)", 2000), hdrCell("الوصف والاستخدام", 4360)] }),
      new TableRow({ children: [dataCell("ad_account", 3000, GRAY_BG), dataCell("account_status", 2000, GRAY_BG), dataCell("تنبيه فوري إذا تم تعطيل الحساب الإعلاني", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("campaign", 3000), dataCell("status", 2000), dataCell("تحديث حالة الحملة في واجهة التطبيق فوراً", 4360)] }),
      new TableRow({ children: [dataCell("leadgen", 3000, GRAY_BG), dataCell("leads", 2000, GRAY_BG), dataCell("استقبال بيانات الـ Leads فور قيام المستخدم بملء الفورم لإرسالها للـ CRM", 4360, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// SECTION 12 - LEAD GENERATION & DETAILED CONTROL
// ============================================================
const sec12 = [
  pageBreak(),
  sectionBanner("القسم الثاني عشر: تفاصيل إضافية وتحكم أدق"),
  h1("13. Lead Generation API"),
  body("يتيح فيسبوك سحب بيانات العملاء المحتملين مباشرة."),
  new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3000, 2000, 4360],
    rows: [
      new TableRow({ children: [hdrCell("الحقل", 3000), hdrCell("النوع", 2000), hdrCell("الوصف", 4360)] }),
      new TableRow({ children: [dataCell("id", 3000, GRAY_BG), dataCell("string", 2000, GRAY_BG), dataCell("معرف الـ Lead", 4360, GRAY_BG)] }),
      new TableRow({ children: [dataCell("created_time", 3000), dataCell("datetime", 2000), dataCell("وقت إنشاء الطلب", 4360)] }),
      new TableRow({ children: [dataCell("field_data", 3000, GRAY_BG), dataCell("array", 2000, GRAY_BG), dataCell("البيانات المعبأة (الاسم، الهاتف، الإيميل...)", 4360, GRAY_BG)] }),
    ]
  }),
];

// ============================================================
// ASSEMBLE
// ============================================================
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "subbullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }
        ] },
      { reference: "numbered",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: BLUE_DARK },
        paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: BLUE_MID },
        paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: TEXT_DARK },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [
    // Cover page
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {},
      footers: {},
      children: coverChildren
    },
    // TOC + Main content
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE_MID, space: 6 } },
            children: [
              new TextRun({ text: "Facebook Ads Multi-Platform Manager  |  GDD v1.0", size: 16, color: TEXT_MID, font: "Arial" }),
              new TextRun({ text: "  ·  صفحة ", size: 16, color: TEXT_MID, font: "Arial" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: TEXT_MID, font: "Arial" }),
            ]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: GRAY_BORDER, space: 6 } },
            children: [new TextRun({ text: "وثيقة سرية — للاستخدام الداخلي فقط — جميع الحقوق محفوظة © 2026", size: 16, color: TEXT_MID, font: "Arial" })]
          })]
        })
      },
      children: [
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 240, after: 160 },
          children: [new TextRun({ text: "فهرس المحتويات", bold: true, size: 36, color: BLUE_DARK, font: "Arial" })]
        }),
        new TableOfContents("فهرس المحتويات", {
          hyperlink: true,
          headingStyleRange: "1-3",
        }),
        pageBreak(),
        ...sec1, ...sec2, ...sec3, ...sec4, ...sec5, ...sec6, ...sec7, ...sec8, ...sec9, ...sec10, ...sec11, ...sec12
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("D:\\facebook_ads\\Facebook_Ads_GDD_v1.0.docx", buffer);
  console.log("DONE: Facebook_Ads_GDD_v1.0.docx written successfully");
}).catch(err => { console.error("ERROR:", err); process.exit(1); });