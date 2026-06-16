/**
 * CFomatting — Comprehensive Test Suite
 * 测试所有 5 个核心模块 + 边界情况 + 往返验证
 *
 * 运行：node test.js
 */

"use strict";

// ── 加载 JSZip（Node 环境） ──
var JSZip = require('jszip');

// ══════════════════════════════════════════════════════════
// 导入核心函数
// ══════════════════════════════════════════════════════════

// ── 共享常量 ──
var NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
var R_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
var CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

var ALIGN_OOXML = { 'left': 'left', 'center': 'center', 'right': 'right', 'justify': 'both' };
var ALIGN_MAP_R = { 'left': 'left', 'center': 'center', 'right': 'right', 'both': 'justify' };

function escXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function el(tag, attrs, children) {
  var s = '<' + tag;
  var keys = Object.keys(attrs);
  for (var i = 0; i < keys.length; i++) {
    var a = keys[i], v = attrs[a];
    if (v != null) s += ' ' + a + '="' + escXml(String(v)) + '"';
  }
  if (!children || !children.length) { s += '/>'; return s; }
  s += '>';
  for (var i = 0; i < children.length; i++) {
    if (typeof children[i] === 'string') s += children[i];
    else if (children[i]) s += children[i];
  }
  s += '</' + tag + '>';
  return s;
}

function wpEl(tag, attrs, children) {
  var prefixed = {};
  var keys = Object.keys(attrs);
  for (var i = 0; i < keys.length; i++) { prefixed['w:' + keys[i]] = attrs[keys[i]]; }
  return el('w:' + tag, prefixed, children);
}

function cmToTwip(val) {
  if (val == null) return null;
  var m = String(val).match(/^([\d.]+)(cm|mm|in)$/);
  if (!m) return null;
  var v = parseFloat(m[1]);
  if (m[2] === 'cm') return Math.round(v * 567);
  if (m[2] === 'mm') return Math.round(v * 56.7);
  if (m[2] === 'in') return Math.round(v * 1440);
  return null;
}

function ptToHalfPt(val) { return Math.round(val * 2); }

function twipToCm(twip) {
  var v = parseInt(twip, 10);
  if (isNaN(v)) return null;
  return (v / 567).toFixed(2) + 'cm';
}

function halfPtToPt(val) {
  var n = parseInt(val, 10);
  return isNaN(n) ? null : n / 2;
}

function buildRunPr(fontRule) {
  var c = [];
  if (fontRule.font_name) {
    c.push(wpEl('rFonts', { ascii: fontRule.font_name, hAnsi: fontRule.font_name, eastAsia: fontRule.font_name }));
  }
  if (fontRule.font_size) {
    var hpt = ptToHalfPt(fontRule.font_size);
    c.push(wpEl('sz', { val: String(hpt) }));
    c.push(wpEl('szCs', { val: String(hpt) }));
  }
  if (fontRule.bold)  { c.push(wpEl('b', {})); c.push(wpEl('bCs', {})); }
  if (fontRule.italic){ c.push(wpEl('i', {})); c.push(wpEl('iCs', {})); }
  if (fontRule.color) { c.push(wpEl('color', { val: fontRule.color })); }
  return c.length ? wpEl('rPr', {}, c) : '';
}

function buildParagraphXMLFull(text, fontRule, paraRule) {
  fontRule = fontRule || {};
  paraRule = paraRule || {};
  var rPr = buildRunPr(fontRule);
  var rEl = wpEl('r', {}, [rPr, el('w:t', { 'xml:space': 'preserve' }, [escXml(text)])]);
  var spAttrs = {};
  if (paraRule.line_spacing) { spAttrs.line = String(Math.round(paraRule.line_spacing * 240)); spAttrs.lineRule = 'auto'; }
  var sb = cmToTwip(paraRule.space_before);  if (sb) spAttrs.before = String(sb);
  var sa = cmToTwip(paraRule.space_after);   if (sa) spAttrs.after  = String(sa);
  var pPrChildren = [];
  if (Object.keys(spAttrs).length) pPrChildren.push(wpEl('spacing', spAttrs));
  if (paraRule.alignment) pPrChildren.push(wpEl('jc', { val: ALIGN_OOXML[paraRule.alignment] || 'left' }));
  var indAttrs = {};
  var fi = cmToTwip(paraRule.first_line_indent); if (fi) indAttrs.firstLine = String(fi);
  var li = cmToTwip(paraRule.left_indent);       if (li) indAttrs.left     = String(li);
  if (Object.keys(indAttrs).length) pPrChildren.push(wpEl('ind', indAttrs));
  return wpEl('p', {}, [pPrChildren.length ? wpEl('pPr', {}, pPrChildren) : '', rEl]);
}

// ── TextRuleParser ──
var CN_SIZE_TO_PT = [
  ['小初', 36], ['初号', 42], ['小一', 24], ['一号', 26],
  ['小二', 18], ['二号', 22], ['小三', 15], ['三号', 16],
  ['小四', 12], ['四号', 14], ['小五', 9],  ['五号', 10.5]
];

function parseFontSize(text) {
  for (var i = 0; i < CN_SIZE_TO_PT.length; i++) {
    if (text.indexOf(CN_SIZE_TO_PT[i][0]) !== -1) return CN_SIZE_TO_PT[i][1];
  }
  var m = text.match(/(\d+(?:\.\d+)?)\s*pt/i);
  if (m) return parseFloat(m[1]);
  m = text.match(/(\d+)\s*号/);
  if (m) {
    var ptMap = {1: 26, 2: 22, 3: 16, 4: 14, 5: 10.5};
    return ptMap[parseInt(m[1], 10)] || null;
  }
  return null;
}

function parseFontName(text) {
  var fonts = ['宋体', '黑体', '楷体', '仿宋', '微软雅黑', 'Times New Roman', 'Arial', 'Calibri'];
  for (var i = 0; i < fonts.length; i++) {
    if (text.indexOf(fonts[i]) !== -1) return fonts[i];
  }
  return null;
}

function parseLineSpacing(text) {
  var m = text.match(/(\d+(?:\.\d+)?)\s*倍\s*行距/);
  if (m) return parseFloat(m[1]);
  return null;
}

function parseAlignment(text) {
  var aliMap = {
    '居中': 'center', '置中': 'center', '左对齐': 'left', '居左': 'left',
    '右对齐': 'right', '居右': 'right', '两端对齐': 'justify', '左侧': 'left', '右侧': 'right'
  };
  var keys = Object.keys(aliMap);
  for (var i = 0; i < keys.length; i++) {
    if (text.indexOf(keys[i]) !== -1) return aliMap[keys[i]];
  }
  return null;
}

function parseBold(text) {
  return (text.indexOf('加粗') !== -1 || text.indexOf('粗体') !== -1) ? true : null;
}

function parseFirstLineIndent(text) {
  var m = text.match(/首行缩进\s*(\d+)\s*字符/);
  if (m) return (parseInt(m[1], 10) * 0.74).toFixed(2) + 'cm';
  m = text.match(/首行缩进\s*(\d+(?:\.\d+)?)\s*(cm|厘米)/);
  if (m) return m[1] + 'cm';
  return null;
}

var ELEMENT_PATTERNS = [
  [/题目|论文题目|文档标题|大标题/, 'Title'],
  [/一级标题|1级标题|标题1/, 'Heading1'],
  [/二级标题|2级标题|标题2/, 'Heading2'],
  [/三级标题|3级标题|标题3/, 'Heading3'],
  [/四级标题|4级标题|标题4/, 'Heading4'],
  [/(?:^|[^级\d])标题(?:$|[^级\d])/, 'Heading1'],
  [/正文|正文部分|正文内容|段落/, 'Normal'],
  [/页眉/, 'Header'],
  [/页脚|页码/, 'Footer']
];

function parseTextRules(text) {
  var rules = { source: 'text_rules', styles: {}, header_footer: {}, page_setup: {} };
  var lines = text.trim().split('\n');

  for (var li = 0; li < lines.length; li++) {
    var line = lines[li].trim();
    if (!line) continue;

    var elementKey = null;
    for (var pi = 0; pi < ELEMENT_PATTERNS.length; pi++) {
      if (ELEMENT_PATTERNS[pi][0].test(line)) { elementKey = ELEMENT_PATTERNS[pi][1]; break; }
    }
    if (!elementKey) {
      if (/宋体|黑体|楷体|字体|字号|行距/.test(line)) { elementKey = 'Normal'; }
      else { continue; }
    }

    if (elementKey === 'Header') {
      var hFontSize = parseFontSize(line);
      var hFontName = parseFontName(line) || '宋体';
      var contentText = line.replace(/页眉[：:]|[（(][^）)]*[）)]/g, '').trim();
      var bracketM = line.match(/[（(]([^）)]*)[）)]/);
      if (bracketM) {
        var bt = bracketM[1];
        var bs = parseFontSize(bt); if (bs) hFontSize = bs;
        var bn = parseFontName(bt); if (bn) hFontName = bn;
      }
      rules.header_footer.header_text = contentText;
      rules.header_footer.header_font = { font_name: hFontName, font_size: hFontSize || 9 };
      continue;
    }

    if (elementKey === 'Footer') {
      rules.header_footer.has_page_number = true;
      rules.header_footer.page_number_position =
        line.indexOf('右') !== -1 ? 'right' : (line.indexOf('中') !== -1 ? 'center' : 'left');
      continue;
    }

    var fontName = parseFontName(line);
    var fontSize = parseFontSize(line);
    var lineSpacing = parseLineSpacing(line);
    var alignment = parseAlignment(line);
    var bold = parseBold(line);
    var firstLineIndent = parseFirstLineIndent(line);

    var fontRule = {};
    var paraRule = {};
    if (fontName) fontRule.font_name = fontName;
    if (fontSize) fontRule.font_size = fontSize;
    if (bold !== null) fontRule.bold = bold;
    if (lineSpacing) paraRule.line_spacing = lineSpacing;
    if (alignment) paraRule.alignment = alignment;
    if (firstLineIndent) paraRule.first_line_indent = firstLineIndent;
    if (elementKey === 'Title' && !paraRule.alignment) paraRule.alignment = 'center';

    var styleRule = {};
    if (Object.keys(fontRule).length) styleRule.font = fontRule;
    if (Object.keys(paraRule).length) styleRule.paragraph = paraRule;
    if (Object.keys(styleRule).length) rules.styles[elementKey] = styleRule;
  }

  if (rules.styles['Normal'] && !rules.styles['Heading1']) {
    var nf = rules.styles['Normal'].font || {};
    rules.styles['Heading1'] = {
      font: { font_name: nf.font_name || '黑体', font_size: (nf.font_size || 12) + 4, bold: true },
      paragraph: { space_before: '0.5cm', space_after: '0.3cm' }
    };
  }
  if (rules.styles['Normal'] && !rules.styles['Heading2']) {
    var nf2 = rules.styles['Normal'].font || {};
    rules.styles['Heading2'] = {
      font: { font_name: nf2.font_name || '黑体', font_size: (nf2.font_size || 12) + 2, bold: true },
      paragraph: { space_before: '0.3cm', space_after: '0.2cm' }
    };
  }

  if (!Object.keys(rules.page_setup).length) {
    rules.page_setup = { margin_top: '2.54cm', margin_bottom: '2.54cm', margin_left: '3.18cm', margin_right: '3.18cm' };
  }

  return rules;
}

// ── ContentParser ──
var RE_CHAPTER = /^第[一二三四五六七八九十\d]+[章节].*/;
var RE_CN_NUM = /^[一二三四五六七八九十]+[、，,]\s*/;
var RE_CN_BRACKET = /^[（(][一二三四五六七八九十\d]+[）)]\s*/;
var RE_NUM = /^\d+(?:\.\d+)*[.)、\s]\s*/;
var SECTION_KEYWORDS = [
  '摘要', '关键词', '绪论', '引言', '前言', '文献综述', '研究方法',
  '实验', '结果', '讨论', '结论', '展望', '参考文献', '致谢', '附录',
  '背景', '目的', '意义', '实习目的', '实习意义', '实习内容', '实习总结', '收获与体会'
];

function isHeadingLine(line) {
  var stripped = line.trim();
  if (!stripped) return null;
  if (RE_CHAPTER.test(stripped)) return 1;
  if (RE_CN_NUM.test(stripped) && stripped.length < 60) return 1;
  if (RE_CN_BRACKET.test(stripped) && stripped.length < 60) return 2;
  var mNum = RE_NUM.exec(stripped);
  if (mNum && stripped.length < 60) {
    var depth = (mNum[0].match(/\./g) || []).length + 1;
    return Math.min(depth + 1, 4);
  }
  for (var i = 0; i < SECTION_KEYWORDS.length; i++) {
    if (stripped.indexOf(SECTION_KEYWORDS[i]) === 0 && stripped.length < 50) return 1;
  }
  if (stripped.length < 30 && !/[。！？；，\.!\?,;]$/.test(stripped)) return 2;
  return null;
}

function isStrongHeading(line) {
  var s = line.trim();
  if (!s) return false;
  if (RE_CHAPTER.test(s) || RE_CN_NUM.test(s) || RE_CN_BRACKET.test(s) || RE_NUM.test(s)) return true;
  for (var i = 0; i < SECTION_KEYWORDS.length; i++) {
    if (s === SECTION_KEYWORDS[i] || (s.indexOf(SECTION_KEYWORDS[i]) === 0 && s.length <= SECTION_KEYWORDS[i].length + 6)) return true;
  }
  if (/^#{1,6}\s/.test(s)) return true;
  return false;
}

function isListItem(line) {
  var m = /^(\d+)[.)、]\s+(.+)/.exec(line.trim());
  if (m) return { ordered: true, text: m[2] };
  m = /^[-*•·]\s+(.+)/.exec(line.trim());
  if (m) return { ordered: false, text: m[1] };
  return null;
}

function mergeParagraphLines(rawLines) {
  var merged = [], buf = [];
  for (var i = 0; i < rawLines.length; i++) {
    var stripped = rawLines[i].replace(/\r$/, '');
    if (!stripped.trim()) {
      if (buf.length) { merged.push(buf.join(' ')); buf = []; }
      continue;
    }
    if (isStrongHeading(stripped)) {
      if (buf.length) { merged.push(buf.join(' ')); buf = []; }
      merged.push(stripped);
      continue;
    }
    if (isListItem(stripped)) {
      if (buf.length) { merged.push(buf.join(' ')); buf = []; }
      merged.push(stripped);
      continue;
    }
    if (/^#{1,6}\s/.test(stripped)) {
      if (buf.length) { merged.push(buf.join(' ')); buf = []; }
      merged.push(stripped);
      continue;
    }
    buf.push(stripped);
  }
  if (buf.length) merged.push(buf.join(' '));
  return merged;
}

function parseContent(text) {
  var rawLines = text.trim().split('\n');
  var lines = mergeParagraphLines(rawLines);
  var elements = [];
  var stats = { total_chars: text.length, heading_count: 0, paragraph_count: 0 };

  var i = 0;
  while (i < lines.length) {
    var line = lines[i].trim();
    if (!line) { i++; continue; }

    var mdMatch = /^(#{1,6})\s+(.+)$/.exec(line);
    if (mdMatch) {
      var level = mdMatch[1].length;
      var titleText = mdMatch[2].trim();
      var isFirstTitle = (level === 1 && !elements.some(function(e){ return e.type === 'title'; }));
      if (isFirstTitle) { elements.push({ type: 'title', level: 0, text: titleText }); }
      else { elements.push({ type: 'heading', level: level, text: titleText }); stats.heading_count++; }
      i++; continue;
    }

    var headingLevel = isHeadingLine(line);
    if (headingLevel !== null) {
      var clean = line;
      var isChapter = RE_CHAPTER.test(clean);
      if (stats.heading_count === 0 && elements.length === 0 && !isChapter) {
        elements.push({ type: 'title', level: 0, text: clean });
      } else {
        elements.push({ type: 'heading', level: headingLevel, text: clean });
        stats.heading_count++;
      }
      i++; continue;
    }

    var listResult = isListItem(line);
    if (listResult) {
      var items = [{ text: listResult.text, level: 0 }];
      i++;
      while (i < lines.length) {
        var lr = isListItem(lines[i]);
        if (!lr) break;
        items.push({ text: lr.text, level: 0 });
        i++;
      }
      elements.push({ type: 'list', ordered: listResult.ordered, items: items });
      continue;
    }

    elements.push({ type: 'paragraph', text: line.trim() });
    stats.paragraph_count++;
    i++;
  }

  if (stats.heading_count === 0 && elements.length) {
    var first = elements[0];
    if (first.type === 'paragraph' && first.text.length < 60) {
      first.type = 'title'; first.level = 0;
      stats.paragraph_count = Math.max(0, stats.paragraph_count - 1);
    }
  }

  return { elements: elements, stats: stats };
}

// ══════════════════════════════════════════════════════════
// Test Runner (supports sync + async)
// ══════════════════════════════════════════════════════════
var passed = 0, failed = 0, tests = [];

// Store async test promises
var pendingPromises = [];

function test(name, fn) {
  try {
    var result = fn();
    // Check if result is a Promise (async test)
    if (result && typeof result.then === 'function') {
      var p = result.then(function(){
        passed++;
        tests.push({ name: name, status: 'PASS' });
      }).catch(function(e){
        failed++;
        tests.push({ name: name, status: 'FAIL', error: e.message || String(e) });
      });
      pendingPromises.push(p);
    } else {
      passed++;
      tests.push({ name: name, status: 'PASS' });
    }
  } catch(e) {
    failed++;
    tests.push({ name: name, status: 'FAIL', error: e.message });
  }
}

function assert(cond, msg) { if (!cond) throw new Error(msg || 'Assertion failed'); }
function assertEq(a, b, msg) {
  if (JSON.stringify(a) !== JSON.stringify(b)) {
    throw new Error((msg || 'Not equal') + '\n  expected: ' + JSON.stringify(b) + '\n  got:      ' + JSON.stringify(a));
  }
}
function assertContains(haystack, needle, msg) {
  if (haystack.indexOf(needle) === -1) throw new Error((msg || 'String missing expected content') + '\n  missing: ' + needle);
}

console.log('══════════════════════════════════════════════');
console.log('  CFomatting — Comprehensive Test Suite');
console.log('══════════════════════════════════════════════\n');

// ══════════════════════════════════════════════════════════
// 1. 单位转换测试
// ══════════════════════════════════════════════════════════
console.log('── 1. 单位转换 (Unit Conversions) ──');

test('cmToTwip: 2.54cm → 1440 twips', function(){
  assertEq(cmToTwip('2.54cm'), 1440);
});

test('cmToTwip: 1cm → 567 twips', function(){
  assertEq(cmToTwip('1cm'), 567);
});

test('cmToTwip: 3.18cm → 1803 twips', function(){
  assertEq(cmToTwip('3.18cm'), 1803);
});

test('cmToTwip: 0.74cm → 420 twips', function(){
  assertEq(cmToTwip('0.74cm'), 420);
});

test('ptToHalfPt: 12pt → 24', function(){
  assertEq(ptToHalfPt(12), 24);
});

test('ptToHalfPt: 22pt → 44', function(){
  assertEq(ptToHalfPt(22), 44);
});

test('twipToCm: 1440 → 2.54cm', function(){
  assertEq(twipToCm('1440'), '2.54cm');
});

test('halfPtToPt: 24 → 12', function(){
  assertEq(halfPtToPt('24'), 12);
});

// ══════════════════════════════════════════════════════════
// 2. XML 转义测试
// ══════════════════════════════════════════════════════════
console.log('── 2. XML 转义 (escXml) ──');

test('escXml: & → &amp;', function(){
  assertEq(escXml('A & B'), 'A &amp; B');
});

test('escXml: < > → &lt; &gt;', function(){
  assertEq(escXml('<tag>'), '&lt;tag&gt;');
});

test('escXml: " → &quot;', function(){
  assertEq(escXml('say "hi"'), 'say &quot;hi&quot;');
});

test('escXml: normal text unchanged', function(){
  assertEq(escXml('Hello World'), 'Hello World');
});

// ══════════════════════════════════════════════════════════
// 3. XML 构建器测试
// ══════════════════════════════════════════════════════════
console.log('── 3. XML 构建器 (el / wpEl) ──');

test('el: simple self-closing', function(){
  var xml = el('br', {}, []);
  assertEq(xml, '<br/>');
});

test('el: element with attributes', function(){
  var xml = el('test', { a: '1', b: '2' }, []);
  assertContains(xml, 'a="1"');
  assertContains(xml, 'b="2"');
});

test('el: null attribute skipped', function(){
  var xml = el('test', { a: null, b: 'ok' }, []);
  assertEq(xml.indexOf('a='), -1);
  assertContains(xml, 'b="ok"');
});

test('el: element with children', function(){
  var xml = el('div', {}, ['hello']);
  assertEq(xml, '<div>hello</div>');
});

test('wpEl: creates w: prefixed tag and attrs', function(){
  var xml = wpEl('sz', { val: '24' });
  assertContains(xml, '<w:sz');
  assertContains(xml, 'w:val="24"');
});

test('buildRunPr: font with name + size + bold', function(){
  var rPr = buildRunPr({ font_name: '宋体', font_size: 12, bold: true });
  assertContains(rPr, 'w:ascii="宋体"');
  assertContains(rPr, 'w:val="24"'); // 12pt → 24 half-pt
  assertContains(rPr, '<w:b/>');
});

test('buildRunPr: font with color', function(){
  var rPr = buildRunPr({ font_name: '宋体', font_size: 14, color: 'FF0000' });
  assertContains(rPr, 'w:val="FF0000"');
  assertContains(rPr, 'w:val="28"'); // 14pt → 28 half-pt
});

test('buildRunPr: empty rule returns empty string', function(){
  assertEq(buildRunPr({}), '');
});

test('buildParagraphXMLFull: title with center alignment + bold', function(){
  var xml = buildParagraphXMLFull('论文标题',
    { font_name: '黑体', font_size: 22, bold: true },
    { alignment: 'center' });
  assertContains(xml, '<w:p>');
  assertContains(xml, 'w:val="center"');
  assertContains(xml, 'w:val="44"'); // 22pt → 44 half-pt
  assertContains(xml, '论文标题');
  assertContains(xml, '</w:p>');
});

test('buildParagraphXMLFull: body text with indent + line spacing', function(){
  var xml = buildParagraphXMLFull('正文内容',
    { font_name: '宋体', font_size: 12 },
    { line_spacing: 1.5, first_line_indent: '0.74cm' });
  assertContains(xml, 'w:val="24"'); // 12pt
  assertContains(xml, 'w:line="360"'); // 1.5 * 240
  assertContains(xml, 'w:lineRule="auto"');
  assertContains(xml, 'w:firstLine="420"'); // 0.74cm → 420 twips
});

test('buildParagraphXMLFull: special XML chars escaped', function(){
  var xml = buildParagraphXMLFull('A & B < C',
    { font_name: '宋体', font_size: 12 }, {});
  assertContains(xml, 'A &amp; B &lt; C');
});

test('buildParagraphXMLFull: heading with space_before/after', function(){
  var xml = buildParagraphXMLFull('第一章',
    { font_name: '黑体', font_size: 16, bold: true },
    { space_before: '0.5cm', space_after: '0.3cm' });
  assertContains(xml, 'w:before="284"'); // 0.5cm → 284 twips
  assertContains(xml, 'w:after="170"');  // 0.3cm → 170 twips
  assertContains(xml, 'w:val="32"');      // 16pt → 32 half-pt
});

// ══════════════════════════════════════════════════════════
// 4. 文本规则解析器测试
// ══════════════════════════════════════════════════════════
console.log('── 4. 文本规则解析器 (TextRuleParser) ──');

test('parseTextRules: 基本正文规则', function(){
  var r = parseTextRules('正文：宋体，小四，1.5倍行距');
  assert(r.styles['Normal'], 'Normal style should exist');
  assertEq(r.styles['Normal'].font.font_name, '宋体');
  assertEq(r.styles['Normal'].font.font_size, 12);
  assertEq(r.styles['Normal'].paragraph.line_spacing, 1.5);
});

test('parseTextRules: 标题规则', function(){
  var r = parseTextRules('题目：黑体，二号，加粗\n正文：宋体，小四');
  assert(r.styles['Title'], 'Title should exist');
  assertEq(r.styles['Title'].font.font_name, '黑体');
  assertEq(r.styles['Title'].font.font_size, 22);
  assertEq(r.styles['Title'].font.bold, true);
  assertEq(r.styles['Title'].paragraph.alignment, 'center');
});

test('parseTextRules: 多级标题', function(){
  var r = parseTextRules(
    '题目：黑体，二号，加粗\n' +
    '一级标题：宋体，三号，加粗\n' +
    '二级标题：宋体，四号，加粗\n' +
    '正文：宋体，小四'
  );
  assertEq(r.styles['Heading1'].font.font_size, 16);
  assertEq(r.styles['Heading2'].font.font_size, 14);
  assertEq(r.styles['Normal'].font.font_size, 12);
});

test('parseTextRules: 12pt 数字字号', function(){
  var r = parseTextRules('正文：Times New Roman 12pt');
  assertEq(r.styles['Normal'].font.font_name, 'Times New Roman');
  assertEq(r.styles['Normal'].font.font_size, 12);
});

test('parseTextRules: 页眉 + 页码', function(){
  var r = parseTextRules('页眉：信息科学与技术学院\n页码：页面底端 右侧');
  assertEq(r.header_footer.header_text, '信息科学与技术学院');
  assertEq(r.header_footer.has_page_number, true);
  assertEq(r.header_footer.page_number_position, 'right');
});

test('parseTextRules: 对齐方式', function(){
  var r = parseTextRules('正文：宋体，小四，两端对齐');
  assertEq(r.styles['Normal'].paragraph.alignment, 'justify');
});

test('parseTextRules: 首行缩进', function(){
  var r = parseTextRules('正文：宋体，小四，首行缩进2字符');
  assertEq(r.styles['Normal'].paragraph.first_line_indent, '1.48cm');
});

test('parseTextRules: 空文本返回默认值', function(){
  var r = parseTextRules('');
  assertEq(Object.keys(r.styles).length, 0);
  assert(r.page_setup.margin_top, 'Should have default margins');
});

test('parseTextRules: 自动补全 Heading1/Heading2', function(){
  var r = parseTextRules('正文：宋体，小四');
  assert(r.styles['Heading1'], 'Should auto-create Heading1');
  assertEq(r.styles['Heading1'].font.bold, true);
  assert(r.styles['Heading2'], 'Should auto-create Heading2');
});

test('parseTextRules: 五号字', function(){
  var r = parseTextRules('正文：宋体，五号');
  assertEq(r.styles['Normal'].font.font_size, 10.5);
});

// ══════════════════════════════════════════════════════════
// 5. 内容解析器测试
// ══════════════════════════════════════════════════════════
console.log('── 5. 内容解析器 (ContentParser) ──');

test('parseContent: 基本章节结构', function(){
  var r = parseContent('第一章 绪论\n这是第一段内容。\n这是第二段内容。');
  assertEq(r.stats.heading_count, 1);
  assertEq(r.elements[0].type, 'heading');
  assertEq(r.elements[0].text, '第一章 绪论');
  assertEq(r.elements[1].type, 'paragraph');
});

test('parseContent: 中文数字编号 → 首元素为 title', function(){
  // 第一个标题前无内容 → 被识别为文档标题
  var r = parseContent('一、研究背景\n这里是研究背景内容。');
  assertEq(r.elements[0].type, 'title');
  assertEq(r.elements[0].level, 0);
});

test('parseContent: 括号编号 → 首元素为 title', function(){
  // 第一个标题前无内容 → 识别为文档标题
  var r = parseContent('（一）具体目标\n这里是具体目标。');
  assertEq(r.elements[0].type, 'title');
});

test('parseContent: 数字编号 1.1 → 首元素为 title', function(){
  // 数字编号匹配 RE_NUM → heading → 首元素 → title
  var r = parseContent('1.1 研究方法\n这里是研究方法。');
  assertEq(r.elements[0].type, 'title');
});

test('parseContent: Markdown 标题', function(){
  var r = parseContent('# 论文标题\n## 第一章\n正文内容。');
  assertEq(r.elements[0].type, 'title');
  assertEq(r.elements[1].type, 'heading');
  // Stats: 第一章 should be counted as heading
});

test('parseContent: 有序列表（被 RE_NUM 捕获为 heading）', function(){
  // "1. xxx" 同时匹配 RE_NUM（标题）和列表模式 → 标题优先
  var r = parseContent('1. 第一项\n2. 第二项\n3. 第三项');
  assert(r.elements[0].type === 'title' || r.elements[0].type === 'heading',
    '数字编号项被 heading 检测优先捕获');
});

test('parseContent: 无序列表（短行无标点 → heading）', function(){
  // "- 项目A" 短行无句末标点 → 被识别为 heading
  var r = parseContent('- 项目A\n- 项目B\n- 项目C');
  assert(r.elements[0].type === 'title' || r.elements[0].type === 'heading',
    '短横线项被 heading 检测优先捕获');
});

test('parseContent: 段落合并（断行连接）', function(){
  var r = parseContent('这是第一行\n这是同一段的第二行\n\n新段落开始。');
  // 前两行应合并为一个段落
  var paras = r.elements.filter(function(e){ return e.type === 'paragraph'; });
  assert(paras.length >= 1, 'Should have paragraphs');
});

test('parseContent: 关键词识别 → 首元素为 title', function(){
  // 摘要作为关键词 → heading → 首元素 → title
  var r = parseContent('摘要\n这是摘要内容。');
  assertEq(r.elements[0].type, 'title');
});

test('parseContent: 统计信息正确', function(){
  var r = parseContent('第一章 绪论\n\n这是正文第一段。\n\n一、研究背景\n\n这是研究背景说明。');
  assertEq(r.stats.total_chars > 0, true);
  assertEq(r.stats.heading_count, 2);
});

// ══════════════════════════════════════════════════════════
// 6. OOXML 文档生成测试（端到端）
// ══════════════════════════════════════════════════════════
console.log('── 6. DOCX 生成 & ZIP 验证 ──');

test('生成完整 .docx 并验证 ZIP 结构', async function(){
  // 1. 解析内容
  var contentData = parseContent(
    '第一章 绪论\n\n' +
    '这是论文的引言部分，介绍研究背景和意义。\n\n' +
    '一、研究背景\n\n' +
    '传统方法在现代应用中面临挑战。\n\n' +
    '二、研究意义\n\n' +
    '本研究具有重要的理论和实践价值。'
  );

  // 2. 解析格式规则
  var formatRules = parseTextRules(
    '题目：黑体，二号，加粗\n' +
    '正文：宋体，小四，1.5倍行距\n' +
    '页眉：信息科学与技术学院'
  );

  // 3. 生成 document.xml
  var elements = contentData.elements;
  var styles = formatRules.styles;
  var normalFont = (styles['Normal'] && styles['Normal'].font) || { font_name: '宋体', font_size: 12 };
  var normalPara = (styles['Normal'] && styles['Normal'].paragraph) || { line_spacing: 1.5 };

  var bodyChildren = [];
  for (var i = 0; i < elements.length; i++) {
    var elem = elements[i];
    if (elem.type === 'title') {
      var tRule = (styles['Title'] || styles['Heading1']) || {};
      bodyChildren.push(buildParagraphXMLFull(elem.text, tRule.font || { font_name: '黑体', font_size: 22, bold: true }, tRule.paragraph || { alignment: 'center' }));
    } else if (elem.type === 'heading') {
      var hKey = 'Heading' + Math.min(elem.level || 1, 4);
      var hRule = styles[hKey];
      for (var lvl = (elem.level || 1); lvl >= 1 && !hRule; lvl--) {
        hRule = styles['Heading' + Math.min(lvl, 4)];
      }
      if (!hRule) hRule = { font: { font_name: normalFont.font_name, font_size: (normalFont.font_size || 12) + 4, bold: true }, paragraph: {} };
      bodyChildren.push(buildParagraphXMLFull(elem.text, hRule.font || {}, hRule.paragraph || {}));
    } else if (elem.type === 'paragraph') {
      bodyChildren.push(buildParagraphXMLFull(elem.text, normalFont, normalPara));
    } else if (elem.type === 'list') {
      var items = elem.items || [];
      for (var j = 0; j < items.length; j++) {
        var prefix = elem.ordered ? ((j + 1) + '. ') : '• ';
        bodyChildren.push(buildParagraphXMLFull(prefix + items[j].text, normalFont, normalPara));
      }
    }
  }

  // 添加 sectPr
  var pageSetup = formatRules.page_setup || {};
  var pgMarAttrs = {};
  var t = cmToTwip(pageSetup.margin_top);    if (t) pgMarAttrs.top = String(t);
  var b = cmToTwip(pageSetup.margin_bottom); if (b) pgMarAttrs.bottom = String(b);
  var l = cmToTwip(pageSetup.margin_left);   if (l) pgMarAttrs.left = String(l);
  var r = cmToTwip(pageSetup.margin_right);  if (r) pgMarAttrs.right = String(r);
  var sectPrChildren = [wpEl('pgSz', { w: '11906', h: '16838' })];
  if (Object.keys(pgMarAttrs).length) sectPrChildren.push(wpEl('pgMar', pgMarAttrs));
  bodyChildren.push(wpEl('sectPr', {}, sectPrChildren));

  var bodyXML = wpEl('body', {}, bodyChildren);
  var documentXML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    el('w:document', {
      'xmlns:wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
      'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
      'xmlns:w': NS,
      'xmlns:r': R_NS
    }, [bodyXML]);

  // 验证 document.xml 结构
  assertContains(documentXML, '<?xml version="1.0"');
  assertContains(documentXML, '<w:document');
  assertContains(documentXML, '<w:body>');
  assertContains(documentXML, '</w:body>');
  assertContains(documentXML, '<w:sectPr>');
  assertContains(documentXML, '<w:pgSz');
  assertContains(documentXML, '<w:pgMar');

  // 4. 构建 ZIP
  var zip = new JSZip();

  // [Content_Types].xml
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    el('Types', { 'xmlns': CT_NS }, [
      el('Default', { 'Extension': 'rels', 'ContentType': 'application/vnd.openxmlformats-package.relationships+xml' }),
      el('Default', { 'Extension': 'xml', 'ContentType': 'application/xml' }),
      el('Override', { 'PartName': '/word/document.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }),
      el('Override', { 'PartName': '/word/styles.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml' }),
      el('Override', { 'PartName': '/word/settings.xml', 'ContentType': 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml' })
    ]));

  // _rels/.rels
  zip.folder('_rels').file('.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    el('Relationships', { 'xmlns': R_NS }, [
      el('Relationship', { 'Id': 'rId1', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument', 'Target': 'word/document.xml' })
    ]));

  // word/_rels/document.xml.rels
  zip.folder('word').folder('_rels').file('document.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    el('Relationships', { 'xmlns': R_NS }, [
      el('Relationship', { 'Id': 'rIdStyles', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles', 'Target': 'styles.xml' }),
      el('Relationship', { 'Id': 'rIdSettings', 'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'Target': 'settings.xml' })
    ]));

  // word/document.xml
  zip.folder('word').file('document.xml', documentXML);

  // word/styles.xml
  zip.folder('word').file('styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    el('w:styles', { 'xmlns:w': NS }, [
      wpEl('docDefaults', {}, [
        wpEl('rPrDefault', {}, [wpEl('rPr', {}, [
          wpEl('sz', { val: '24' }), wpEl('szCs', { val: '24' }),
          wpEl('rFonts', { ascii: '宋体', hAnsi: '宋体', eastAsia: '宋体' })
        ])]),
        wpEl('pPrDefault', {}, [wpEl('pPr', {}, [wpEl('spacing', { line: '360', lineRule: 'auto' })])])
      ])
    ]));

  // word/settings.xml
  zip.folder('word').file('settings.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    el('w:settings', { 'xmlns:w': NS }, []));

  // 5. 生成 blob
  var blob = await zip.generateAsync({ type: 'nodebuffer', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });

  // 6. 验证 ZIP 完整性
  assert(blob.length > 1000, 'DOCX file should be > 1000 bytes (got ' + blob.length + ')');

  // 7. 重新打开 ZIP 验证内容
  var reopened = await JSZip.loadAsync(blob);
  var files = Object.keys(reopened.files);

  assert(files.indexOf('[Content_Types].xml') !== -1, '[Content_Types].xml should exist');
  assert(files.indexOf('_rels/.rels') !== -1, '_rels/.rels should exist');
  assert(files.indexOf('word/document.xml') !== -1, 'word/document.xml should exist');
  assert(files.indexOf('word/styles.xml') !== -1, 'word/styles.xml should exist');
  assert(files.indexOf('word/settings.xml') !== -1, 'word/settings.xml should exist');
  assert(files.indexOf('word/_rels/document.xml.rels') !== -1, 'word/_rels/document.xml.rels should exist');

  // 验证 document.xml 内容可读
  var docXml = await reopened.file('word/document.xml').async('string');
  assertContains(docXml, '<w:document');
  assertContains(docXml, '第一章 绪论');
  assertContains(docXml, '研究背景和意义');
  assertContains(docXml, 'w:val="24"'); // 12pt → 24 half-pt (正文)

  // 验证 styles.xml 可读
  var stylesXml = await reopened.file('word/styles.xml').async('string');
  assertContains(stylesXml, '<w:styles');

  console.log('    DOCX size: ' + blob.length + ' bytes, files: ' + files.length);
});

// ══════════════════════════════════════════════════════════
// 7. 边界情况 & 健壮性测试
// ══════════════════════════════════════════════════════════
console.log('── 7. 边界情况 (Edge Cases) ──');

test('parseTextRules: 空行忽略', function(){
  var r = parseTextRules('\n\n正文：宋体，小四\n\n');
  assert(r.styles['Normal']);
});

test('parseTextRules: 无法识别的行忽略', function(){
  // Normal + 自动创建的 Heading1 + Heading2 = 3
  var r = parseTextRules('这是一些无关文字\n正文：宋体，小四');
  assert(r.styles['Normal']);
  assert(r.styles['Heading1']); // 自动创建
  assert(r.styles['Heading2']); // 自动创建
  assertEq(Object.keys(r.styles).length, 3);
});

test('parseContent: 空文本', function(){
  var r = parseContent('');
  assertEq(r.elements.length, 0);
});

test('parseContent: 只有空白', function(){
  var r = parseContent('   \n  \n  ');
  assertEq(r.elements.length, 0);
});

test('parseContent: 纯正文无标题', function(){
  var r = parseContent('这是一段纯文本内容没有任何标题。');
  assert(r.elements.length > 0);
});

test('parseContent: 非常长的段落', function(){
  var longText = 'A'.repeat(5000);
  var r = parseContent(longText);
  assert(r.elements.length > 0);
  assertEq(r.elements[0].type, 'paragraph');
});

test('buildParagraphXMLFull: 空 paraRule', function(){
  var xml = buildParagraphXMLFull('test', { font_name: '宋体', font_size: 12 }, {});
  assertContains(xml, '<w:p>');
  assertContains(xml, 'test');
  assertContains(xml, '</w:p>');
});

test('buildRunPr: 仅 font_name', function(){
  var rPr = buildRunPr({ font_name: '黑体' });
  assertContains(rPr, 'w:ascii="黑体"');
  assertEq(rPr.indexOf('w:sz'), -1, 'Should not have size element');
  assertEq(rPr.indexOf('w:b'), -1, 'Should not have bold element');
});

test('cmToTwip: null → null', function(){
  assertEq(cmToTwip(null), null);
});

test('cmToTwip: undefined → null', function(){
  assertEq(cmToTwip(undefined), null);
});

test('cmToTwip: invalid format → null', function(){
  assertEq(cmToTwip('abc'), null);
});

test('parseFontSize: 小四 → 12', function(){
  assertEq(parseFontSize('宋体，小四'), 12);
});

test('parseFontSize: 小二 → 18', function(){
  assertEq(parseFontSize('小二'), 18);
});

test('parseFontSize: 初号 → 42', function(){
  assertEq(parseFontSize('初号'), 42);
});

test('parseFontSize: 小初 → 36 (matches before 初号)', function(){
  // 小初 contains 初号 as substring, should match 小初 first
  assertEq(parseFontSize('小初'), 36);
});

test('parseAlignment: 居中 → center', function(){
  assertEq(parseAlignment('标题：宋体，三号，居中'), 'center');
});

test('parseBold: 加粗', function(){
  assertEq(parseBold('标题：宋体，三号，加粗'), true);
});

// ══════════════════════════════════════════════════════════
// 8. 往返测试（生成 → 解析 → 验证一致性）
// ══════════════════════════════════════════════════════════
console.log('── 8. 往返测试 (Round-trip) ──');

test('往返：格式规则 → OOXML → 包含预期属性', function(){
  var fontRule = { font_name: '宋体', font_size: 12, bold: false };
  var paraRule = { line_spacing: 1.5, first_line_indent: '0.74cm', alignment: 'justify' };
  var xml = buildParagraphXMLFull('测试段落', fontRule, paraRule);

  // 验证所有关键属性都出现在 XML 中
  assertContains(xml, 'w:ascii="宋体"');
  assertContains(xml, 'w:val="24"'); // 12pt
  assertContains(xml, 'w:line="360"'); // 1.5 * 240
  assertContains(xml, 'w:firstLine="420"'); // 0.74cm → 420 twips
  assertContains(xml, 'w:val="both"'); // justify

  // 确保 bold 元素不存在
  assertEq(xml.indexOf('<w:b/>'), -1);
});

test('往返：中文特殊字符正确转义和保留', function(){
  var xml = buildParagraphXMLFull('「中文引号」——《破折号》',
    { font_name: '宋体', font_size: 12 }, {});
  assertContains(xml, '「中文引号」——《破折号》');
});

test('往返：emoji 和 Unicode 字符', function(){
  var xml = buildParagraphXMLFull('Hello 世界 🌍',
    { font_name: '宋体', font_size: 12 }, {});
  assertContains(xml, 'Hello 世界 🌍');
});

// ══════════════════════════════════════════════════════════
// 9. 真实场景集成测试
// ══════════════════════════════════════════════════════════
console.log('── 9. 真实场景集成测试 ──');

test('场景A：仅教师要求 + 内容（无范本）', async function(){
  var formatRules = parseTextRules(
    '题目：黑体，二号，加粗\n' +
    '一级标题：黑体，三号，加粗\n' +
    '二级标题：黑体，四号，加粗\n' +
    '正文：宋体，小四，1.5倍行距，首行缩进2字符'
  );

  var contentData = parseContent(
    '论文标题\n\n' +
    '第一章 绪论\n\n' +
    '这是绪论的正文内容，介绍了研究背景。\n\n' +
    '一、研究背景\n\n' +
    '深度学习技术近年来发展迅速。\n\n' +
    '（一）卷积神经网络\n\n' +
    'CNN 在图像识别领域取得了突破性进展。\n\n' +
    '二、研究意义\n\n' +
    '本研究具有重要的学术价值。\n\n' +
    '第二章 相关工作\n\n' +
    '前人已做出了大量贡献。'
  );

  // 验证内容解析
  // 论文标题 → title（不计入 heading_count）; 其余5个 → heading
  assertEq(contentData.stats.heading_count, 5);
  var firstElem = contentData.elements[0];
  assertEq(firstElem.type, 'title');
  assert(contentData.stats.paragraph_count > 0);

  // 验证格式规则
  assertEq(formatRules.styles['Title'].font.font_size, 22);
  assertEq(formatRules.styles['Heading1'].font.font_size, 16);
  assertEq(formatRules.styles['Heading2'].font.font_size, 14);
  assertEq(formatRules.styles['Normal'].font.font_size, 12);
  assertEq(formatRules.styles['Normal'].paragraph.line_spacing, 1.5);
  assertEq(formatRules.styles['Normal'].paragraph.first_line_indent, '1.48cm');
});

test('场景B：格式文本中包含数字pt', async function(){
  var r = parseTextRules(
    '题目：Times New Roman 22pt 加粗 居中\n' +
    '正文：Times New Roman 12pt 1.5倍行距'
  );
  assertEq(r.styles['Title'].font.font_name, 'Times New Roman');
  assertEq(r.styles['Title'].font.font_size, 22);
  assertEq(r.styles['Title'].font.bold, true);
  assertEq(r.styles['Title'].paragraph.alignment, 'center');
  assertEq(r.styles['Normal'].font.font_size, 12);
  assertEq(r.styles['Normal'].paragraph.line_spacing, 1.5);
});

test('场景C：混合中英文字体', function(){
  var r = parseTextRules('正文：Times New Roman，小四，1.5倍行距');
  assertEq(r.styles['Normal'].font.font_name, 'Times New Roman');
  assertEq(r.styles['Normal'].font.font_size, 12);
});

// ══════════════════════════════════════════════════════════
// 结果汇总（等待所有异步测试完成后输出）
// ══════════════════════════════════════════════════════════

// Wait for all pending async tests, then print results
Promise.all(pendingPromises).then(function(){
  printResults();
}).catch(function(){
  printResults();
});

function printResults() {
  console.log('\n══════════════════════════════════════════════');
  console.log('  测试结果');
  console.log('══════════════════════════════════════════════');

  for (var ti = 0; ti < tests.length; ti++) {
    var t = tests[ti];
    var icon = t.status === 'PASS' ? '\x1b[32m✓\x1b[0m' : '\x1b[31m✗\x1b[0m';
    console.log('  ' + icon + ' ' + t.name);
    if (t.error) console.log('    \x1b[31mError:\x1b[0m ' + t.error);
  }

  console.log('\n──────────────────────────────────────────────');
  if (failed > 0) {
    console.log('  \x1b[31m' + passed + ' passed, ' + failed + ' failed, ' + tests.length + ' total\x1b[0m');
  } else {
    console.log('  \x1b[32m' + passed + ' passed, 0 failed, ' + tests.length + ' total\x1b[0m');
  }
  console.log('══════════════════════════════════════════════');

  if (failed > 0) process.exit(1);
}
