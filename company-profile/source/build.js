// Builds the redesigned Sama'a company profile deck (Arabic, RTL).
const fs = require('fs');
const path = require('path');
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const React = require('react');
const ReactDOMServer = require('react-dom/server');
const Tb = require('react-icons/tb');

const A = (p) => path.join(__dirname, 'assets', p);
const OUT = process.argv[2] || 'Samaa_Company_Profile.pptx';

// ---- Brand palette (sampled from the original deck) ----
const C = {
  bg: '1E2831',
  card: '26333E',
  cardLine: '364752',
  teal: '128A9E',
  tealBright: '2AAFC4',
  purple: '662E8F',
  beige: 'C2BFAF',
  ink: 'F1EFE8',
  body: 'B9C1C7',
  white: 'FFFFFF',
};
const F = {
  title: 'FF Taweel',
  bold: 'Almarai Bold',
  light: 'Almarai Light',
  xbold: 'Almarai ExtraBold',
};
const W = 13.333, H = 7.5;

// ---- helpers ----
function txt(slide, text, o) {
  slide.addText(text, Object.assign({
    isTextBox: true, rtlMode: true, lang: 'ar-SA', margin: 0,
    align: 'right', valign: 'top', fontFace: F.light, color: C.body, fontSize: 14,
  }, o));
}

async function icon(name, color, px = 256) {
  const Comp = Tb[name];
  if (!Comp) throw new Error('missing icon ' + name);
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(Comp, { color: '#' + color, size: px, strokeWidth: 1.5 }));
  const buf = await sharp(Buffer.from(svg)).resize(px, px).png().toBuffer();
  return 'image/png;base64,' + buf.toString('base64');
}

function imgData(file) {
  const ext = path.extname(file).slice(1).replace('jpg', 'jpeg');
  return `image/${ext};base64,` + fs.readFileSync(file).toString('base64');
}

async function logoFit(file, bw, bh) {
  const m = await sharp(file).metadata();
  const r = Math.min(bw / m.width, bh / m.height);
  return { w: m.width * r, h: m.height * r };
}

// shared chrome for content slides: small logo mark + brand band
function chrome(slide, bg = 'bg_clean.jpg') {
  slide.background = { color: C.bg };
  if (bg) slide.addImage({ data: imgData(A(bg)), x: 0, y: 0, w: W, h: H });
  slide.addImage({ data: imgData(A('image8_t.png')), x: 0.3, y: 0.28, w: 0.34, h: 0.34 / 0.6245 });
  slide.addShape('rect', { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: C.teal }, line: { type: 'none' } });
}

// section heading block, anchored on the right (RTL start)
function heading(slide, title, subtitle, y = 0.55) {
  slide.addImage({ data: imgData(A('image6_t.png')), x: W - 0.78, y: y + 0.08, w: 0.26, h: 0.26 / 0.3139 });
  txt(slide, title, { x: 5.2, y, w: W - 1.05 - 5.2, h: 0.95, fontFace: F.title, fontSize: 48, color: C.beige, valign: 'middle' });
  if (subtitle) txt(slide, subtitle, { x: 4.2, y: y + 1.0, w: W - 1.05 - 4.2, h: 0.4, fontFace: F.bold, fontSize: 15, color: C.tealBright, valign: 'middle' });
}

function iconRing(slide, x, y, d, data, fill = C.bg) {
  slide.addShape('ellipse', { x, y, w: d, h: d, fill: { color: fill }, line: { color: C.teal, width: 1.5 } });
  const p = d * 0.25;
  slide.addImage({ data, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
}

function card(slide, x, y, w, h, fill = C.card) {
  slide.addShape('roundRect', {
    x, y, w, h, rectRadius: 0.14, fill: { color: fill }, line: { color: C.cardLine, width: 0.75 },
    shadow: { type: 'outer', color: '000000', blur: 12, offset: 3, angle: 90, opacity: 0.35 },
  });
}

(async () => {
  const pres = new pptxgen();
  pres.layout = 'LAYOUT_WIDE';
  pres.title = 'سمع للإنتاج الإعلامي - الملف التعريفي';
  pres.company = "Sama'a Audiovisual Productions";
  pres.rtlMode = true;
  pres.theme = { headFontFace: F.bold, bodyFontFace: F.light, lang: 'ar-SA' };

  const ICON = C.ink;
  const icons = {};
  for (const n of ['TbFeather', 'TbMusic', 'TbAdjustmentsHorizontal', 'TbMicrophone', 'TbSpeakerphone', 'TbDisc',
    'TbBroadcast', 'TbHeadphones', 'TbPhoneCall', 'TbCut', 'TbUsers', 'TbMovie', 'TbMicrophone2', 'TbWaveSine', 'TbVinyl']) {
    icons[n] = await icon(n, ICON);
  }

  // ============ 1. Cover (kept as the original) ============
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    s.addImage({ data: imgData(A('bg_rings.jpg')), x: 0, y: 0, w: W, h: H });
    const lw = 4.68, lh = lw / (5430 / 4426);
    s.addImage({ data: imgData(A('image2_t.png')), x: (W - lw) / 2, y: (H - lh) / 2 - 0.1, w: lw, h: lh });
    s.addNotes('غلاف الملف التعريفي لشركة سمع للإنتاج الإعلامي المرئي والمسموع.');
  }

  // ============ 2. About us ============
  {
    const s = pres.addSlide();
    chrome(s);
    heading(s, 'من نحن', 'سمع للإنتاج الإعلامي المرئي والمسموع', 0.7);
    txt(s, [
      { text: 'شركة سمع للإنتاج الإعلامي شركة متخصصة في إنتاج المحتوى الغنائي بأشكاله المتعددة، عن طريق نخبة من المتخصصين ذوي الخبرة الواسعة في المجال، وباستخدام أحدث تقنيات الإنتاج المرئي والمسموع.', options: { breakLine: true, paraSpaceAfter: 10 } },
      { text: 'كما نقدم خدمات الإنتاج الصوتي المحترف بخبرة تتجاوز 25 عاماً؛ من البودكاست والتعليق الصوتي، إلى التصميم الصوتي ورسائل الرد الآلي للشركات.' },
    ], { x: 6.3, y: 2.55, w: W - 1.05 - 6.3, h: 2.5, fontSize: 16, color: C.ink, lineSpacingMultiple: 1.25 });

    // two mini stats under the paragraph
    const stats = [
      { n: '13', l: 'خدمة إبداعية متخصصة' },
      { n: '2', l: 'لغتان: العربية والإنجليزية' },
    ];
    const sw = 2.85, sx0 = W - 1.05 - sw;
    stats.forEach((st, i) => {
      const x = sx0 - i * (sw + 0.3), y = 5.35;
      card(s, x, y, sw, 1.15);
      s.addText(st.n, { isTextBox: true, x: x + sw - 1.05, y: y + 0.12, w: 0.85, h: 0.9, margin: 0, align: 'center', valign: 'middle', fontFace: F.xbold, fontSize: 36, color: C.tealBright });
      txt(s, st.l, { x: x + 0.2, y: y + 0.12, w: sw - 1.35, h: 0.9, fontFace: F.bold, fontSize: 13, color: C.ink, valign: 'middle' });
    });

    // big purple "25+" disc on the left
    const d = 3.7, cx = 0.95, cy = 1.75;
    s.addShape('ellipse', { x: cx - 0.28, y: cy - 0.28, w: d + 0.56, h: d + 0.56, fill: { type: 'none' }, line: { color: C.teal, width: 1 } });
    s.addShape('ellipse', { x: cx, y: cy, w: d, h: d, fill: { color: C.purple }, line: { type: 'none' } });
    s.addText('+25', { isTextBox: true, x: cx, y: cy + 0.75, w: d, h: 1.4, margin: 0, align: 'center', valign: 'middle', fontFace: F.xbold, fontSize: 80, color: C.white });
    txt(s, 'عاماً من الخبرة', { x: cx, y: cy + 2.15, w: d, h: 0.5, align: 'center', fontFace: F.bold, fontSize: 20, color: C.beige });
    s.addNotes('نبذة عن الشركة مع أبرز الأرقام: أكثر من 25 عاماً من الخبرة، 13 خدمة متخصصة، ولغتان.');
  }

  // ============ 3. Clients ============
  {
    const s = pres.addSlide();
    chrome(s);
    heading(s, 'شركاء النجاح', 'نفخر بخدمة نخبة من أبرز العلامات التجارية');
    const order = ['c03_image11', 'c11_image19', 'c09_image17', 'c00_image7', 'c10_image18', 'c15_image23',
      'c14_image22', 'c16_image24', 'c05_image13', 'c13_image21', 'c12_image20', 'c07_image15',
      'c02_image10', 'c08_image16', 'c04_image12', 'c01_image9', 'c17_image25', 'c06_image14'];
    const cols = 6, tw = 1.85, th = 1.3, gx = 0.2, gy = 0.22;
    const x0 = (W - (cols * tw + (cols - 1) * gx)) / 2, y0 = 2.3;
    // logos that carry their own dark background get a matching tile
    const tileFill = { c05_image13: '010101', c04_image12: '343434' };
    for (let i = 0; i < order.length; i++) {
      const col = cols - 1 - (i % cols), row = Math.floor(i / cols); // RTL fill
      const x = x0 + col * (tw + gx), y = y0 + row * (th + gy);
      const f = A(`clients/${order[i]}_trim.png`);
      const fill = tileFill[order[i]] || 'FFFFFF';
      s.addShape('roundRect', { x, y, w: tw, h: th, rectRadius: 0.12, fill: { color: fill }, line: { type: 'none' },
        shadow: { type: 'outer', color: '000000', blur: 10, offset: 3, angle: 90, opacity: 0.35 } });
      const pad = fill === 'FFFFFF' ? 0.2 : 0.08;
      const fit = await logoFit(f, tw - 2 * pad, th - 2 * pad);
      s.addImage({ data: imgData(f), x: x + (tw - fit.w) / 2, y: y + (th - fit.h) / 2, w: fit.w, h: fit.h });
    }
    s.addNotes('شعارات بعض عملاء الشركة وشركاء النجاح.');
  }

  // ============ 4. Services divider ============
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    s.addImage({ data: imgData(A('bg_rings.jpg')), x: 0, y: 0, w: W, h: H });
    s.addImage({ data: imgData(A('image8_t.png')), x: 0.3, y: 0.28, w: 0.34, h: 0.34 / 0.6245 });
    s.addShape('rect', { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: C.teal }, line: { type: 'none' } });
    txt(s, 'خدماتنا', { x: 2, y: 0.85, w: W - 4, h: 1.4, align: 'center', valign: 'middle', fontFace: F.title, fontSize: 72, color: C.beige });
    txt(s, 'حلول إبداعية متكاملة في الإنتاج الموسيقي والصوتي', { x: 2, y: 2.3, w: W - 4, h: 0.5, align: 'center', valign: 'middle', fontFace: F.bold, fontSize: 18, color: C.tealBright });

    const pillars = [
      { n: '01', t: 'الإنتاج الموسيقي والغنائي', d: 'كتابة الكلمات • تأليف الألحان • التوزيع الموسيقي • التسجيل • الأغاني الدعائية • الألبومات', ic: 'TbVinyl' },
      { n: '02', t: 'الإنتاج الصوتي', d: 'البودكاست • الرد الآلي • المعلقون الصوتيون • تحرير الصوت • التصميم الصوتي • التعليق الصوتي', ic: 'TbWaveSine' },
    ];
    const pw = 5.0, ph = 2.75, gap = 0.5, px0 = (W - (2 * pw + gap)) / 2, py = 3.45;
    pillars.forEach((p, i) => {
      const x = px0 + (1 - i) * (pw + gap);
      card(s, x, py, pw, ph);
      iconRing(s, x + pw - 0.35 - 0.95, py + 0.35, 0.95, icons[p.ic], C.purple);
      s.addText(p.n, { isTextBox: true, x: x + 0.35, y: py + 0.35, w: 1.4, h: 0.95, margin: 0, align: 'left', valign: 'middle', fontFace: F.xbold, fontSize: 40, color: '4F6370' });
      txt(s, p.t, { x: x + 0.35, y: py + 1.45, w: pw - 0.7, h: 0.5, fontFace: F.bold, fontSize: 20, color: C.ink, valign: 'middle' });
      txt(s, p.d, { x: x + 0.35, y: py + 1.95, w: pw - 0.7, h: 0.65, fontSize: 12, color: C.body, lineSpacingMultiple: 1.2 });
    });
    s.addNotes('محورا خدمات الشركة: الإنتاج الموسيقي والغنائي، والإنتاج الصوتي.');
  }

  // ============ 5. Music services (original six, refined) ============
  {
    const s = pres.addSlide();
    chrome(s, null);
    // signature purple disc carrying the title, as in the original
    const d = 5.0, cx = 9.25, cy = 1.15;
    s.addShape('ellipse', { x: cx - 0.3, y: cy - 0.3, w: d + 0.6, h: d + 0.6, fill: { type: 'none' }, line: { color: C.teal, width: 1 } });
    s.addShape('ellipse', { x: cx, y: cy, w: d, h: d, fill: { color: C.purple }, line: { type: 'none' } });
    txt(s, 'خدماتنا', { x: cx + 0.35, y: cy + 1.35, w: W - cx - 0.7, h: 1.3, align: 'center', valign: 'middle', fontFace: F.title, fontSize: 60, color: C.beige });
    txt(s, 'الإنتاج الموسيقي والغنائي', { x: cx + 0.35, y: cy + 2.65, w: W - cx - 0.7, h: 0.5, align: 'center', valign: 'middle', fontFace: F.bold, fontSize: 17, color: C.white });

    const items = [
      { ic: 'TbFeather', t: 'كتابة الكلمات', d: 'شبكة واسعة من الشعراء والكتّاب ذوي الخبرة في كتابة الكلمات الغنائية والشعر العربي، بإبداع وسرعة في الإنجاز.' },
      { ic: 'TbMusic', t: 'تأليف الألحان', d: 'من أكثر خدماتنا تميزاً؛ ألحان شرقية وغربية على أيدي ملحنين محترفين عاشقين للفن بكل أنواعه وتفاصيله.' },
      { ic: 'TbAdjustmentsHorizontal', t: 'التوزيع الموسيقي', d: 'الورقة الرابحة لعملائنا دوماً؛ توزيع موسيقي دقيق وعلى مستوى فني احترافي عالٍ في استديوهات سمع.' },
      { ic: 'TbMicrophone', t: 'التسجيل', d: 'خدمات التسجيل بجميع أنواعها: أغانٍ دعائية، ألبومات، فويس أوفر، والعديد من الخدمات الأخرى.' },
      { ic: 'TbSpeakerphone', t: 'أغانٍ دعائية', d: 'صناعة المحتوى الغنائي الدعائي للشركات والجهات المعلنة لتحقيق أفضل النتائج التسويقية.' },
      { ic: 'TbDisc', t: 'أغاني الألبومات', d: 'منظومة متكاملة لإنتاج ألبومات مطربي الصف الأول في الخليج والعالم العربي، ودعم الفنانين الصاعدين.' },
    ];
    const colW = 3.95, gapX = 0.35, rowH = 2.0, x0 = 0.55, y0 = 1.0, ring = 0.8;
    items.forEach((it, i) => {
      const col = 1 - (i % 2), row = Math.floor(i / 2);
      const x = x0 + col * (colW + gapX), y = y0 + row * rowH;
      iconRing(s, x + colW - ring, y, ring, icons[it.ic]);
      txt(s, it.t, { x, y: y + 0.05, w: colW - ring - 0.25, h: 0.45, fontFace: F.bold, fontSize: 17, color: C.ink, valign: 'middle' });
      txt(s, it.d, { x, y: y + 0.55, w: colW - ring - 0.25, h: 1.15, fontSize: 12, color: C.body, lineSpacingMultiple: 1.2 });
    });
    s.addNotes('خدمات الإنتاج الموسيقي والغنائي الست.');
  }

  // ============ 6. NEW: Audio production services ============
  {
    const s = pres.addSlide();
    chrome(s);
    heading(s, 'الإنتاج الصوتي', 'خدمات إنتاج صوتي احترافية بخبرة تتجاوز 25 عاماً');

    const items = [
      null, // highlight card
      { ic: 'TbBroadcast', t: 'إنتاج البودكاست الصوتي', d: 'إنتاج متكامل لحلقات البودكاست باحترافية حتى النسخة النهائية الجاهزة للنشر.' },
      { ic: 'TbHeadphones', t: 'هندسة وإخراج البودكاست', d: 'للبودكاست المسجَّل من طرفكم: هندسة صوتية وإخراج فني بجودة احترافية.' },
      { ic: 'TbPhoneCall', t: 'رسائل الرد الآلي للشركات', d: 'رسائل ترحيب وقوائم رد آلي بأصوات احترافية تعكس هوية شركتكم.' },
      { ic: 'TbCut', t: 'تحرير الصوت للمعلقين الصوتيين', d: 'تنقية وتحرير ومعالجة تسجيلات المعلقين لتسليمها بجودة عالية.' },
      { ic: 'TbUsers', t: 'توفير معلقين صوتيين', d: 'معلقون ومعلقات باللغتين العربية والإنجليزية لتناسب كل مشروع.' },
      { ic: 'TbMovie', t: 'التصميم الصوتي لمقاطع الفيديو', d: 'مؤثرات وموسيقى ومكساج يمنح مقاطع الفيديو حضوراً وتأثيراً أقوى.' },
      { ic: 'TbMicrophone2', t: 'التعليق الصوتي للأفلام الوثائقية والموشن جرافيك', d: 'تعليق صوتي احترافي ينقل رسالتكم بوضوح وإحساس.' },
    ];
    const cols = 4, cw = 2.88, ch = 2.3, gx = 0.25, gy = 0.25;
    const x0 = (W - (cols * cw + (cols - 1) * gx)) / 2, y0 = 2.1;
    items.forEach((it, i) => {
      const col = cols - 1 - (i % cols), row = Math.floor(i / cols);
      const x = x0 + col * (cw + gx), y = y0 + row * (ch + gy);
      if (!it) {
        s.addShape('roundRect', { x, y, w: cw, h: ch, rectRadius: 0.14, fill: { color: C.purple }, line: { type: 'none' },
          shadow: { type: 'outer', color: '000000', blur: 12, offset: 3, angle: 90, opacity: 0.35 } });
        s.addText('+25', { isTextBox: true, x: x + 0.25, y: y + 0.3, w: cw - 0.5, h: 0.95, margin: 0, align: 'right', valign: 'middle', fontFace: F.xbold, fontSize: 54, color: C.white });
        txt(s, 'عاماً من الخبرة', { x: x + 0.25, y: y + 1.3, w: cw - 0.5, h: 0.4, fontFace: F.bold, fontSize: 17, color: C.white, valign: 'middle' });
        txt(s, 'في الإنتاج الصوتي المحترف', { x: x + 0.25, y: y + 1.7, w: cw - 0.5, h: 0.35, fontSize: 12, color: 'E4D6EE', valign: 'middle' });
        return;
      }
      card(s, x, y, cw, ch);
      const ring = 0.72;
      iconRing(s, x + cw - 0.25 - ring, y + 0.25, ring, icons[it.ic]);
      txt(s, it.t, { x: x + 0.25, y: y + 1.07, w: cw - 0.5, h: 0.55, fontFace: F.bold, fontSize: 13, color: C.ink, valign: 'top', lineSpacingMultiple: 1.05 });
      txt(s, it.d, { x: x + 0.25, y: y + 1.62, w: cw - 0.5, h: 0.6, fontSize: 10.5, color: C.body, lineSpacingMultiple: 1.15 });
    });
    s.addNotes('خدمات الإنتاج الصوتي الجديدة: البودكاست، هندسة وإخراج البودكاست، الرد الآلي، تحرير الصوت، المعلقون الصوتيون، التصميم الصوتي، والتعليق الصوتي.');
  }

  // ============ 7. Thank you ============
  {
    const s = pres.addSlide();
    s.background = { color: C.bg };
    s.addImage({ data: imgData(A('bg_rings.jpg')), x: 0, y: 0, w: W, h: H });
    s.addImage({ data: imgData(A('image8_t.png')), x: 0.3, y: 0.28, w: 0.34, h: 0.34 / 0.6245 });
    s.addShape('rect', { x: 0, y: H - 0.28, w: W, h: 0.28, fill: { color: C.teal }, line: { type: 'none' } });
    const mw = 1.05, mh = mw / 0.6245;
    s.addImage({ data: imgData(A('image8_t.png')), x: (W - mw) / 2, y: 1.2, w: mw, h: mh });
    txt(s, 'شكراً لكم', { x: 2.5, y: 3.15, w: W - 5, h: 1.5, align: 'center', valign: 'middle', fontFace: F.title, fontSize: 80, color: C.beige });
    txt(s, 'سمع للإنتاج الإعلامي المرئي والمسموع', { x: 2.5, y: 4.75, w: W - 5, h: 0.5, align: 'center', valign: 'middle', fontFace: F.bold, fontSize: 18, color: C.tealBright });
    s.addImage({ data: imgData(A('image6_t.png')), x: 4.05, y: 3.3, w: 0.26, h: 0.26 / 0.3139 });
    s.addImage({ data: imgData(A('image6_t.png')), x: 9.0, y: 2.6, w: 0.26, h: 0.26 / 0.3139 });
  }

  // pptxgenjs tags every run with East Asian (ea) and complex-script (cs) fonts
  // using Chinese code pages (-122 GB2312, -120 Big5). Without the brand fonts
  // installed, PowerPoint then falls back to a Chinese font and mobile apps draw
  // the Arabic as empty squares. Rewrite each run to the original deck's
  // Arabic-only setup: latin + cs with the Arabic code page (-78), no ea.
  const JSZip = require('jszip');
  const zip = await JSZip.loadAsync(await pres.write({ outputType: 'nodebuffer' }));
  const fontRun = /<a:latin typeface="([^"]*)"[^>]*\/><a:ea typeface="[^"]*"[^>]*\/><a:cs typeface="[^"]*"[^>]*\/>/g;
  for (const name of Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    const xml = (await zip.file(name).async('string'))
      .replace(fontRun, '<a:latin typeface="$1" pitchFamily="2" charset="-78"/><a:cs typeface="$1" pitchFamily="2" charset="-78"/>')
      .replace(/ altLang="en-US"/g, '')
      .replace(/ lang="en-US"/g, ' lang="ar-SA"');
    if (/<a:ea |charset="-12[02]"/.test(xml)) throw new Error('unpatched East Asian font in ' + name);
    zip.file(name, xml);
  }
  fs.writeFileSync(OUT, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  console.log('wrote', OUT);
})();
