/* 报销单系统  v2.9 — 2026-09-17
 * 纯本地网页：发票上传 -> 多格式识别 -> 自动填入「费用报销单」
 *
 * 报销单版式严格对齐 报销单_54_fee.xlsx「费用报销单」表样（B:N 共 13 列）：
 *   行1 标题(合并13列)   行2 部门(B:C) | 报销日期(D:N)
 *   行3-4 表头 摘要(B:D 跨2行) | 金额(E:L) + 八位分格行 | 科目(M 跨2行) | 单据张数(N 跨2行)
 *   行5-9 明细 摘要(B:D) | 八位分格(E:L 十万 万 千 百 十 元 角 分，前导位留空) | 科目(M) | 张数(N)
 *   行10 合计（大写）(B:D) | 八位分格 | 单据合计(M:N)
 *   行11 签字栏 财会主管(B) 记账(C) 经办人(D) 部门主管(E:J) 复核标签(K) 复核输入框(L:M=分+科目) 报销人(N)；各栏均可手填
 *
 * 引擎：① 优先用随包 lib/（用启动.bat 以 http 打开时完全离线）② 否则按国内镜像链自动降级
 * 数据：只存在本机浏览器内存，不上传任何服务器。
 */
(function () {
'use strict';

/* ==================================================================
 * 一、引擎加载
 * ================================================================== */
var LOCAL_BASE = 'lib/';
var IS_HTTP = (location.protocol === 'http:' || location.protocol === 'https:');

var MIRRORS = [
  { id: '本地随包', type: 'local' },
  { id: 'jsDelivr', type: 'jsd', base: 'https://cdn.jsdelivr.net/npm/' },
  { id: 'jsDelivr镜像1', type: 'jsd', base: 'https://fastly.jsdelivr.net/npm/' },
  { id: 'jsDelivr镜像2', type: 'jsd', base: 'https://gcore.jsdelivr.net/npm/' },
  { id: 'unpkg', type: 'jsd', base: 'https://unpkg.com/' },
  { id: 'npmmirror', type: 'npm', base: 'https://registry.npmmirror.com/' }
];

var LIBS = {
  tesseract: {
    label: '图片识别', local: 'tesseract/tesseract.min.js',
    jsd: 'tesseract.js@5.1.1/dist/tesseract.min.js',
    npm: 'tesseract.js/5.1.1/files/dist/tesseract.min.js',
    test: function () { return !!(window.Tesseract && window.Tesseract.recognize); }
  },
  pdfjs: {
    label: 'PDF解析', local: 'pdf.min.js',
    jsd: 'pdfjs-dist@3.11.174/build/pdf.min.js',
    npm: 'pdfjs-dist/3.11.174/files/build/pdf.min.js',
    test: function () { return !!(window.pdfjsLib && window.pdfjsLib.getDocument); }
  },
  xlsx: {
    label: 'Excel解析', local: 'xlsx.full.min.js',
    jsd: 'xlsx@0.18.5/dist/xlsx.full.min.js',
    npm: 'xlsx/0.18.5/files/dist/xlsx.full.min.js',
    test: function () { return !!(window.XLSX && window.XLSX.utils); }
  },
  /* 导出用：能写「真样式」（边框/字体/合并/列宽/行高）的 xlsx 生成库。
     随包 lib/ 里是隔离了全局名（XLSX_STYLE）的版本，避免与上面的解析库打架；
     走 CDN 时拿不到该全局名，会自动回退到旧的 HTML .xls 方式。 */
  xlsxStyle: {
    label: 'Excel导出', local: 'xlsx.style.min.js',
    jsd: 'xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
    npm: 'xlsx-js-style/1.2.0/files/dist/xlsx.bundle.js',
    test: function () { return !!(window.XLSX_STYLE && window.XLSX_STYLE.utils && window.XLSX_STYLE.write); }
  },
  mammoth: {
    label: 'Word解析', local: 'mammoth.browser.min.js',
    jsd: 'mammoth@1.8.0/mammoth.browser.min.js',
    npm: 'mammoth/1.8.0/files/mammoth.browser.min.js',
    test: function () { return !!window.mammoth; }
  },
  html2canvas: {
    label: '导出出图', local: 'html2canvas.min.js',
    jsd: 'html2canvas@1.4.1/dist/html2canvas.min.js',
    npm: 'html2canvas/1.4.1/files/dist/html2canvas.min.js',
    test: function () { return !!window.html2canvas; }
  },
  jspdf: {
    label: 'PDF导出', local: 'jspdf.umd.min.js',
    jsd: 'jspdf@2.5.1/dist/jspdf.umd.min.js',
    npm: 'jspdf/2.5.1/files/dist/jspdf.umd.min.js',
    test: function () { return !!((window.jspdf && window.jspdf.jsPDF) || window.jsPDF); }
  },
  jsqr: {
    label: '二维码', local: 'jsqr.min.js',
    jsd: 'jsqr@1.4.0/dist/jsQR.min.js',
    npm: 'jsqr/1.4.0/files/dist/jsQR.min.js',
    test: function () { return !!window.jsQR; }
  }
};

var engineState = {};   // name -> {status:'pending'|'ok'|'fail', mirror, note}
var loadPromises = {};
Object.keys(LIBS).forEach(function (k) { engineState[k] = { status: 'pending', mirror: '', note: '' }; });

var ocrBase = null;      // {workerPath, corePath, langPath}
var pdfWorkerUrl = '';

function urlFor(m, item) {
  if (m.type === 'local') return LOCAL_BASE + item.local;
  if (m.type === 'jsd') return m.base + item.jsd;
  return m.base + item.npm;
}

function injectScript(url) {
  return new Promise(function (res, rej) {
    var s = document.createElement('script');
    s.async = true;
    var done = false;
    var timer = setTimeout(function () {
      if (done) return; done = true;
      rej(new Error('加载超时'));
    }, 25000);
    s.onload = function () { if (done) return; done = true; clearTimeout(timer); res(); };
    s.onerror = function () {
      if (done) return; done = true; clearTimeout(timer);
      if (s.parentNode) s.parentNode.removeChild(s);
      rej(new Error('加载失败'));
    };
    s.src = url;
    document.head.appendChild(s);
  });
}

function setOcrBase(m) {
  if (m.type === 'local') {
    ocrBase = {
      workerPath: LOCAL_BASE + 'tesseract/worker.min.js',
      corePath: LOCAL_BASE + 'tesseract-core',
      langPath: LOCAL_BASE + 'tessdata'
    };
  } else if (m.type === 'jsd') {
    ocrBase = {
      workerPath: m.base + 'tesseract.js@5.1.1/dist/worker.min.js',
      corePath: m.base + 'tesseract.js-core@5.1.1',
      langPath: m.base + '@tesseract.js-data/chi_sim@1.0.0/4.0.0_best_int'
    };
  } else {
    ocrBase = {
      workerPath: m.base + 'tesseract.js/5.1.1/files/dist/worker.min.js',
      corePath: m.base + 'tesseract.js-core/5.1.1/files',
      langPath: m.base + '@tesseract.js-data/chi_sim/1.0.0/files/4.0.0_best_int'
    };
  }
}
function setPdfWorker(m) {
  if (m.type === 'local') pdfWorkerUrl = LOCAL_BASE + 'pdf.worker.min.js';
  else if (m.type === 'jsd') pdfWorkerUrl = m.base + 'pdfjs-dist@3.11.174/build/pdf.worker.min.js';
  else pdfWorkerUrl = m.base + 'pdfjs-dist/3.11.174/files/build/pdf.worker.min.js';
}

function loadLib(name) {
  if (loadPromises[name]) return loadPromises[name];
  var item = LIBS[name];
  var cands = MIRRORS.filter(function (m) { return m.type !== 'local' || IS_HTTP; });
  var chain = cands.map(function (m) { return { m: m, url: urlFor(m, item) }; });
  var p = (function () {
    var i = 0;
    function step() {
      if (i >= chain.length) {
        engineState[name] = { status: 'fail', mirror: '', note: '所有镜像都连不上' };
        updateEngineStatus();
        return Promise.resolve(false);
      }
      var cur = chain[i++];
      return injectScript(cur.url).then(function () {
        if (item.test()) {
          engineState[name] = { status: 'ok', mirror: cur.m.id, note: '' };
          if (name === 'tesseract') setOcrBase(cur.m);
          if (name === 'pdfjs') setPdfWorker(cur.m);
          updateEngineStatus();
          return true;
        }
        return step();
      }, function () { return step(); });
    }
    return step();
  })();
  loadPromises[name] = p;
  return p;
}
function reloadEngines() {
  loadPromises = {};
  Object.keys(LIBS).forEach(function (k) { engineState[k] = { status: 'pending', mirror: '', note: '' }; });
  updateEngineStatus();
  Object.keys(LIBS).forEach(function (k) { loadLib(k); });
}

function updateEngineStatus() {
  var box = $('#engineStatus');
  if (!box) return;
  var html = '';
  Object.keys(LIBS).forEach(function (k) {
    var st = engineState[k];
    var cls = st.status === 'ok' ? 'es-ok' : (st.status === 'fail' ? 'es-bad' : 'es-wait');
    var txt = st.status === 'ok' ? '就绪' : (st.status === 'fail' ? '不可用' : '加载中');
    html += '<span class="es ' + cls + '" title="' + esc(LIBS[k].label + '｜' + txt + '｜' + (st.mirror || st.note)) + '">' +
      LIBS[k].label + '·' + txt + '</span>';
  });
  box.innerHTML = html;
  var srcs = Object.keys(LIBS).filter(function (k) { return engineState[k].status === 'ok'; })
    .map(function (k) { return engineState[k].mirror; });
  var uniq = srcs.filter(function (v, i) { return srcs.indexOf(v) === i; });
  var foot = $('#engineSource');
  if (!foot) {
    foot = document.createElement('div');
    foot.id = 'engineSource';
    foot.className = 'es-src';
    box.parentNode.appendChild(foot);
  }
  foot.textContent = uniq.length ? ('来源：' + uniq.join('、') + (IS_HTTP ? '' : '（直接双击打开时无法读本地引擎）')) : '正在连接…';
  /* 识别组件（OCR / 二维码）都不可用：提示用户可直接手动填写，不影响使用 */
  var mh = $('#manualHint');
  if (mh) {
    var ocrOk = engineState.tesseract && engineState.tesseract.status === 'ok';
    var qrOk = engineState.jsqr && engineState.jsqr.status === 'ok';
    mh.classList.toggle('hidden', ocrOk || qrOk);
  }
}

/* ==================================================================
 * 二、基础工具
 * ================================================================== */
function $(s) { return document.querySelector(s); }
function $$(s) { return Array.prototype.slice.call(document.querySelectorAll(s)); }
var seq = 1;
function uid() { return 'i' + (seq++) + Date.now().toString(36); }
function pad(n) { n = String(n); return n.length < 2 ? '0' + n : n; }
function cnDate(d) { d = d || new Date(); return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日'; }
/* 把发票日期统一成「2026年9月14日」展示（二维码给 YYYY-MM-DD，OCR 给中文，皆兼容） */
function normDate(s) {
  if (!s) return '';
  var m = String(s).trim().match(/^(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})/);
  if (m) return m[1] + '年' + (+m[2]) + '月' + (+m[3]) + '日';
  return String(s).trim();
}
function num(v) { var n = parseFloat(String(v == null ? '' : v).replace(/[,\s¥￥]/g, '')); return isNaN(n) ? 0 : n; }
function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function extOf(name) { var i = name.lastIndexOf('.'); return i < 0 ? '' : name.slice(i + 1).toLowerCase(); }
function baseName(name) { return name.replace(/\.[^.]+$/, ''); }
function flat(s) { return String(s == null ? '' : s).replace(/[\s\u3000\u00a0]+/g, ''); }
function capital2(n) { return (Math.round(num(n) * 100) / 100).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

var toastTimer = null;
function toast(msg) {
  var t = $('#toast'); if (!t) return;
  t.textContent = msg; t.classList.remove('hidden');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { t.classList.add('hidden'); }, 3600);
}

/* 人民币大写 */
function toCapital(money) {
  money = Math.round((Number(money) + 1e-9) * 100) / 100;
  if (money === 0) return '零元整';
  var neg = money < 0; money = Math.abs(money);
  var integer = Math.floor(money);
  var decimal = Math.round((money - integer) * 100);
  var cnNums = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];
  var cnIntUnit = ['', '拾', '佰', '仟'];
  var cnBigUnit = ['', '万', '亿', '兆'];
  function four(num4) {
    var s = '', zero = false, str = String(num4);
    for (var i = 0; i < str.length; i++) {
      var d = +str[i], p = str.length - i - 1;
      if (d === 0) { zero = true; }
      else { if (zero && s !== '') s += '零'; s += cnNums[d] + cnIntUnit[p]; zero = false; }
    }
    return s;
  }
  var intStr = String(integer), groups = [];
  while (intStr.length > 0) { groups.unshift(+intStr.slice(-4)); intStr = intStr.slice(0, -4); }
  var s = '';
  for (var g = 0; g < groups.length; g++) {
    if (groups[g] !== 0) s += four(groups[g]) + cnBigUnit[groups.length - 1 - g];
    else if (g < groups.length - 1 && groups[g + 1] !== 0 && s !== '') s += '零';
  }
  s += '元';
  if (decimal > 0) {
    var j = Math.floor(decimal / 10), f = decimal % 10;
    if (j > 0) s += cnNums[j] + '角';
    if (f > 0) s += cnNums[f] + '分';
  } else s += '整';
  return neg ? '负' + s : s;
}

/* 金额 <-> 八位分格（E:L = 十万 万 千 百 十 元 角 分）；模板最高位「十万」，上限 999999.99 */
var DIGIT_W = [100000, 10000, 1000, 100, 10, 1, 0.1, 0.01];
var DIGIT_LABELS = ['十万', '万', '千', '百', '十', '元', '角', '分'];
var MAX_GRID = 999999.99;
/* 数位格显示：最高有效位之前的位（十万/万）留空不写 0，其余位按位显示（含内部 0、尾随 0）。全部为 0 时全留空。 */
function visibleDigits(ds) {
  ds = ds || [];
  var lead = 0;
  while (lead < ds.length && (ds[lead] === 0 || ds[lead] === '0' || ds[lead] === '' || ds[lead] == null)) lead++;
  var out = [];
  for (var i = 0; i < ds.length; i++) out.push(i < lead ? '' : (ds[i] === 0 || ds[i] === '0' ? '0' : (ds[i] || '')));
  return out;
}
function amountToDigits(n) {
  var cents = Math.round(Math.min(num(n), MAX_GRID) * 100);
  var d = [0, 0, 0, 0, 0, 0, 0, 0];
  for (var i = 7; i >= 0; i--) { d[i] = cents % 10; cents = Math.floor(cents / 10); }
  return d;
}
function digitsToAmount(digits) {
  var s = 0;
  for (var i = 0; i < 8; i++) s += (parseInt(digits[i], 10) || 0) * DIGIT_W[i];
  return Math.round(s * 100) / 100;
}
function sanitizeDigit(v) { v = String(v).replace(/[^0-9]/g, ''); return v.slice(-1); }
function emptyDigits() { return ['', '', '', '', '', '', '', '']; }

/* ==================================================================
 * 三、状态
 * ================================================================== */
function blankRow() {
  return { id: uid(), invId: null, summary: '', amount: 0, digits: emptyDigits(), subject: '', count: '' };
}
var state = {
  dept: '',
  date: '',
  rows: [],
  sign: { caik: '', jizhang: '', jingban: '', bumen: '', fuhe: '', baoxiao: '' },  // 签字栏（可手填）
  invoices: []   // {id, fileName, ext, thumb, status, error, rawText, progress, file}
};
/* 初始给 5 行明细（与原表样 5-9 行一致），每一行都是可直接输入的实体行 */
(function () { for (var i = 0; i < 5; i++) state.rows.push(blankRow()); })();

function computeTotal() { var s = 0; state.rows.forEach(function (r) { s += num(r.amount); }); return s; }
function totalCount() { var c = 0; state.rows.forEach(function (r) { c += num(r.count); }); return Math.round(c); }
function hasData() {
  return state.invoices.length > 0 || state.rows.some(function (r) {
    return r.summary || r.subject || num(r.count) || num(r.amount);
  });
}

/* ==================================================================
 * 三·b、报销台账（localStorage）：已导出/打印的发票记「已报销」，
 *        重复上传时拦截。仅存本机浏览器，换电脑/换浏览器/清数据会失效。
 * ================================================================== */
var REIMB_KEY = 'bx_reimb_v1';
function loadReimb() {
  try { var o = JSON.parse(localStorage.getItem(REIMB_KEY)); if (o && o.hash && o.num) return o; } catch (e) {}
  return { hash: {}, num: {} };
}
function saveReimb(o) { try { localStorage.setItem(REIMB_KEY, JSON.stringify(o)); } catch (e) {} }
function isoToday() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
function fmtLedgerDate(iso) {
  var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? (+m[2]) + '月' + (+m[3]) + '日' : String(iso || '');
}
/* 文件内容指纹：优先 SHA-256（需 https/localhost），否则退回 名称+大小+修改时间 */
function fileHash(file) {
  return readAsArrayBuffer(file).then(function (buf) {
    if (window.crypto && window.crypto.subtle && buf.byteLength < 60 * 1024 * 1024) {
      return window.crypto.subtle.digest('SHA-256', buf).then(function (dig) {
        var a = Array.prototype.slice.call(new Uint8Array(dig));
        return 'sha:' + a.map(function (b) { return ('0' + b.toString(16)).slice(-2); }).join('');
      }).catch(function () { return 'fb:' + file.name + '|' + file.size + '|' + (file.lastModified || 0); });
    }
    return 'fb:' + file.name + '|' + file.size + '|' + (file.lastModified || 0);
  });
}

/* ==================================================================
 * 四、文本读取
 * ================================================================== */
function readAsArrayBuffer(file) {
  return new Promise(function (res, rej) {
    var r = new FileReader();
    r.onload = function () { res(r.result); };
    r.onerror = function () { rej(r.error || new Error('读取失败')); };
    r.readAsArrayBuffer(file);
  });
}
function readAsDataURL(file) {
  return new Promise(function (res, rej) {
    var r = new FileReader();
    r.onload = function () { res(r.result); };
    r.onerror = function () { rej(r.error || new Error('读取失败')); };
    r.readAsDataURL(file);
  });
}

/* 发票内容：增值税发票「货物或应税劳务、服务名称」写作「*类别*名称」
   （如 *生产生活服务*住宿服务、*生产生活服务*代订机票），摘要要的就是第二个 * 之后的名称。
   注意：绝不能退回表头取词——表头「项目名称」右邻是「规格型号」，一旦退回就会把
   「规格型号」当成发票内容写进摘要。 */
var HEADER_WORD = /^(规格型号|规格|型号|单位|数量|单价|金额|税率|征收率|税额|价税合计|合计|备注|项目名称|服务名称|货物或应税劳务|货物或应税劳务、服务名称)$/;
function isHeaderWord(w) { return HEADER_WORD.test(String(w == null ? '' : w).replace(/\s/g, '')); }
function extractGoodsName(s) {
  if (!s) return '';
  var m = String(s).match(/[*＊※]([^*＊※]{1,24})[*＊※]([^*＊※]{1,40})/);
  if (!m) return '';
  var name = m[2].split(/[*＊※]/)[0];                    // 第二个 * 之后、下一个 * 之前
  /* 尾随的数量/金额串（如「住宿服务天3839.62」）从第一个「非数字后跟数字」处截断 */
  var cut = name.match(/^(.*?[^\d０-９.,，．])[\d０-９]/);
  name = cut ? cut[1] : name;
  name = name.replace(/^[^\u4e00-\u9fa5A-Za-z0-9（(]+/, '');
  name = name.replace(/[\s，,。.、：:；;＊*※（）()]+$/g, '').trim();
  if (cut) name = name.replace(/[天日次个台]$/, '').trim();   // 截断后残留的计量单位
  if (name.length > 24) name = name.slice(0, 24);
  return (!name || isHeaderWord(name)) ? '' : name;
}

/* ==================================================================
 * 五、发票字段解析
 * ================================================================== */
function parseInvoice(text) {
  var r = { type: '', code: '', number: '', date: '', seller: '', item: '', total: '', totalGuessed: false };
  var t = flat(text);
  if (!t) return r;

  /* 发票种类 */
  if (/增值税/.test(t) && /专用发票/.test(t)) r.type = /电子/.test(t) ? '电子专票' : '专票';
  else if (/增值税/.test(t) && /普通发票/.test(t)) r.type = /电子/.test(t) ? '电子普票' : '普票';
  else if (/机动车销售统一发票/.test(t)) r.type = '机动车';
  else if (/区块链/.test(t)) r.type = '区块链电子发票';
  else if (/通行费/.test(t)) r.type = '通行费';
  else if (/航空运输电子客票|行程单/.test(t)) r.type = '机票行程单';
  else if (/铁路|火车票/.test(t)) r.type = '火车票';
  else if (/出租车|客运|公路运输/.test(t)) r.type = '客运票';
  else if (/定额/.test(t)) r.type = '定额发票';
  else if (/电子发票/.test(t)) r.type = '电子普票';
  else if (/发票/.test(t)) r.type = '普票';

  var m;
  m = t.match(/发票代码[:：]?(\d{10,12})/); if (m) r.code = m[1];
  m = t.match(/发票号码[:：]?\s*([0-9]{8,20})/); if (m) r.number = m[1];
  if (!r.number) { m = t.match(/票据号码[:：]?\s*([0-9]{8,20})/); if (m) r.number = m[1]; }
  if (!r.number) {
    /* 阅读顺序错位：标签在页首、号码在页尾，中间夹着购买方/销售方信息。
       向后（≤260 字）取首个 10~20 位纯数字串即为发票号码（增值税电子发票号码多为 20 位；
       纳税人识别号含字母或排在其后，不会抢先命中）。二维码优先，此处仅作文字层兜底。 */
    var npos = t.indexOf('发票号码'); if (npos < 0) npos = t.indexOf('票据号码');
    if (npos >= 0) { var nrest = t.slice(npos + 4); var nm = nrest.match(/([0-9]{10,20})/); if (nm) r.number = nm[1]; }
  }
  if (!r.number) { m = t.match(/号码[:：]?\s*([0-9]{8,20})/); if (m) r.number = m[1]; }

  m = t.match(/开票日期[:：]?(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m) r.date = m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  if (!r.date) { m = t.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/); if (m) r.date = m[1] + '-' + pad(m[2]) + '-' + pad(m[3]); }
  if (!r.date) { m = t.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if (m) r.date = m[1] + '-' + pad(m[2]) + '-' + pad(m[3]); }

  /* 价税合计（小写）：电子发票文字层阅读顺序错位时，「（小写）」与其金额被大量
     购买方/销售方文字隔开，故优先按「大写金额 ¥ 数字」定位——大写金额（圆整 / 圆X角X分）
     之后紧跟的 ¥ 两位小数即价税合计，最稳。再退化到就近的「价税合计（小写）¥ 数字」等模式。 */
  m = t.match(/[零壹贰叁肆伍陆柒捌玖拾佰仟万亿]+圆[零壹贰叁肆伍陆柒捌玖拾佰仟万亿角分整零\s]{0,30}?[¥￥]?\s*(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/价税合计[^\d¥￥（(]{0,80}?(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/（小写）[^\d¥￥]{0,12}?[¥￥]\s*(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/小写[^\d¥￥]{0,12}?(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/合计金额[^\d¥￥]{0,20}?(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/\((?:小写)\)[^\d¥￥]{0,12}?(\d[\d,]*\.\d{2})/);
  if (m) r.total = m[1].replace(/,/g, '');
  /* 不再兜底为「全文最大两位小数」：多发票或异常文本会抓到错误大数（如 99999999）。
     取不到价税合计就留空，改由二维码金额或导出时提示手工补填。 */

  /* 销售方公司全称 —— 二维码不含此项，只能靠 OCR，故做多重兜底
     ① 优先「销售方…名称：公司名 纳税人识别号」（兼容 OCR 把「名称」误识为「各称/名你」）
     ② 否则取所有「名称：公司名 纳税人识别号/税号」的最后一处（购买方在前、销售方在后）
     ③ 否则「销方名称：公司名」
     ④ 兜底：扫描含「公司/事务所/集团/店/厂…」的主体名，取「销售方」之后最近的一个 */
  var seller = '';
  m = t.match(/销售方(?:信息)?[一-龥]{0,6}?(?:名你|各称|名称)[:：]?([一-龥A-Za-z0-9（）()·.\-]{2,40}?)(?=统一社会信用代码|纳税人识别号|税号|地址|电话|开户行)/);
  if (!m) m = t.match(/销售方[:：]?([一-龥A-Za-z0-9（）()·.\-]{2,40}?)(?=统一社会信用代码|纳税人识别号|税号|地址|电话|开户行)/);
  if (!m) {
    var all = [], re = /(?:名你|各称|名称)[:：]?([一-龥A-Za-z0-9（）()·.\-]{2,40}?)(?=统一社会信用代码|纳税人识别号|税号)/g, mm;
    while ((mm = re.exec(t)) !== null) all.push(mm[1]);
    if (all.length) m = [null, all[all.length - 1]];
  }
  if (!m) m = t.match(/销方(?:名你|各称|名称)[:：]?([一-龥A-Za-z0-9（）()·.\-]{2,40}?)(?=销方|纳税)/);
  if (m) seller = m[1].replace(/^[：:＊*、，,。.\-]+|[：:＊*、，,。.\-]+$/g, '').trim();
  if (!seller) {
    var si = t.indexOf('销售方'), corps = t.match(/([一-龥A-Za-z0-9（）()·.\-]{2,30}?(?:公司|事务所|集团|店|厂|商行|有限公司|有限责任|中心|学校|医院|银行|合作社))/g) || [];
    for (var ci = 0; ci < corps.length; ci++) {
      var pos = t.indexOf(corps[ci]);
      if (si >= 0 ? pos > si : ci === 0) { seller = corps[ci]; break; }
    }
  }
  if (seller && /购买方|信息/.test(seller)) seller = '';   // 防止抓到「购买方信息」等栏目标签废字
  if (seller && seller.length >= 2) r.seller = seller;

  /* 发票内容：取「*类别*名称」中第二个 * 之后的名称（OCR 文本同样适用） */
  r.item = extractGoodsName(t);
  if (!r.item) {
    /* 少数票把项目名称写得没有星号，则取表头后紧跟的名称——但必须挡掉表头词（规格型号/单位/数量…） */
    m = t.match(/(?:货物或应税劳务[、,，]?服务名称|项目名称)[:：]?\s*([一-龥A-Za-z0-9（）()·]{2,20}?)(?=规格|单位|数量|单价|金额|税率|价税|合计|\s|$)/);
    if (m && !isHeaderWord(m[1])) r.item = m[1].replace(/[，,。.、]/g, '').trim();
  }
  /* 最后兜底：图片发票的星号被 OCR 误识时，按常见费用关键词取一个（不含表头词，不会误取规格型号） */
  if (!r.item) {
    m = t.match(/(服务费|餐饮费|餐费|住宿费|办公用品|交通费|咨询费|培训费|会议费|快递费|物料|手续费|电费|水费|油费|维修费)/);
    if (m) r.item = m[1];
  }
  return r;
}

/* ==================================================================
 * 五·b、电子发票版面感知解析（pdf.js 带坐标的文字层）
 * 发票二维码不含销售方/货物名，摘要只能从文字层取；但 PDF 文字读取顺序
 * 会把左右两栏（购买方/销售方）打乱，故用坐标定位：销售方在右侧栏
 * （x 大于「销售方」标签的 x）。纯本地、不联网、比 OCR 准确得多。
 * ================================================================== */
function parsePdfLayout(items) {
  var res = { seller: '', item: '' };
  if (!items || !items.length) return res;
  var its = [];
  items.forEach(function (it) {
    if (!it || !it.str || !it.str.trim()) return;
    var tr = it.transform || [1, 0, 0, 1, 0, 0];
    its.push({ s: it.str, x: tr[4] || 0, y: tr[5] || 0 });
  });
  if (!its.length) return res;

  /* 销售方标签的位置（排除「购买方」） */
  var sx = null;
  its.forEach(function (it) {
    if (/销售方/.test(it.s) && !/购买方/.test(it.s)) { if (sx === null || it.x > sx) sx = it.x; }
  });

  /* 候选公司名：要求整段文字项本身就是「主体名 + 公司/集团/…」后缀（锚定首尾），
     避免把「银行账号:1051…」这类只含「银行」二字的非公司行误判为公司。
     开户行（银行/支行）虽在右侧但不是销售方，故 seller 优先取「公司类」主体。 */
  var compRe = /^[一-龥A-Za-z0-9（）()·.\-]{2,40}?(?:公司|有限公司|有限责任公司|事务所|集团|酒店|旅行社|商行|中心|学校|医院|合作社|厂|店|超市|科技|实业|工作室)$/;
  var cands = [];
  its.forEach(function (it) {
    if (!compRe.test(it.s)) return;
    if (/账号|税号|识别号|电话|地址|开户行/.test(it.s)) return;   // 排除账号/税号等废行
    var name = it.s.replace(/^(名称|销方|卖方|销售方|购买方)?[:：]?\s*/, '').trim();
    if (name.length >= 4) cands.push({ name: name, x: it.x, y: it.y });
  });

  if (sx !== null) {
    /* 销售方在右侧栏：取 x 大于标签且最靠右的「公司类」主体；无则退而取最靠右候选 */
    var right = cands.filter(function (c) { return c.x > sx - 6; });
    var corp = right.filter(function (c) { return /(公司|集团|事务所|中心|学校|医院|厂|店|商行|合作社|科技|实业|工作室|有限责任公司)$/.test(c.name); });
    if (corp.length) res.seller = corp[corp.length - 1].name;
    else if (right.length) res.seller = right[right.length - 1].name;
  }
  if (!res.seller && cands.length) res.seller = cands[cands.length - 1].name;

  /* 发票内容：取「*类别*名称」第二个 * 之后的内容。
     货物行在 PDF 里通常是独立文字项，先逐项匹配（最准）；不中再把文字项拼起来匹配；
     仍不中才用坐标兜底（「项目名称」表头正下方最左的那段文字）。
     全程不退回“表头邻格”取词，避免把「规格型号」写进摘要。 */
  var all = its.map(function (it) { return it.s; }).join('');
  var got = '';
  for (var gi = 0; gi < its.length && !got; gi++) got = extractGoodsName(its[gi].s);
  if (!got) got = extractGoodsName(all);
  if (!got) got = goodsNameByCoord(its);
  res.item = got;
  return res;
}

/* 坐标兜底：发票「项目名称」表头正下方、最靠左的一段文字即货物或服务名称 */
function goodsNameByCoord(its) {
  var hdr = null;
  its.forEach(function (it) {
    if (/(货物或应税劳务|服务名称|项目名称)/.test(it.s)) { if (!hdr || it.y > hdr.y) hdr = it; }
  });
  if (!hdr) return '';
  var cand = its.filter(function (it) {
    return it.y < hdr.y - 0.5 && it.y > hdr.y - 24 &&
      /[一-龥A-Za-z]/.test(it.s) && !isHeaderWord(it.s);
  });
  if (!cand.length) return '';
  cand.sort(function (a, b) { return a.x - b.x; });
  var nm = extractGoodsName(cand[0].s) || cand[0].s.replace(/^[*＊※\s]+/, '').trim();
  return isHeaderWord(nm) ? '' : nm;
}

/* 专票/普票：优先用二维码里的发票类型（机器读取，最准），其次文字层/OCR 标题里的「增值税专用发票」。
   返回 '专票' / '普票' / ''（完全没有可识别文字时留空，交手工填写）。 */
function detectVat(text, qr) {
  if (qr && qr.type) {
    if (/专用/.test(qr.type)) return '专票';
    if (/普通|通行费/.test(qr.type)) return '普票';
  }
  if (text && /[一-龥]/.test(text)) {
    return /增值税专用发票/.test(text) ? '专票' : '普票';
  }
  return '';
}

/* ==================================================================
 * 五·a、发票二维码识别（国税局规范：01,类型,代码,号码,金额,日期,校验码）
 * 二维码是机器读取，几乎不会错，且金额是完整数值；优先于 OCR 用于填表。
 * ================================================================== */
var QR_TYPE = {
  '01': '增值税专用发票', '04': '增值税普通发票', '08': '电子专票',
  '10': '增值税电子普通发票', '11': '通行费电子普通发票', '12': '收费公路通行费电子普通发票',
  '32': '数电发票'
};

/* 解码整张图里的二维码，返回明文串或 null */
function decodeQr(dataUrl) {
  return loadLib('jsqr').then(function (ok) {
    if (!ok || !window.jsQR) return null;
    return new Promise(function (res) {
      var img = new Image();
      img.onload = function () {
        try {
          var w = img.naturalWidth, h = img.naturalHeight;
          if (!w || !h) { res(null); return; }
          var cap = 2200, scale = Math.min(1, cap / Math.max(w, h));
          var cw = Math.max(1, Math.round(w * scale)), ch = Math.max(1, Math.round(h * scale));
          var c = document.createElement('canvas'); c.width = cw; c.height = ch;
          var ctx = c.getContext('2d');
          ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cw, ch);
          ctx.drawImage(img, 0, 0, cw, ch);
          var tryCanvas = function (cv) {
            try {
              var id = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height);
              var r = window.jsQR(id.data, id.width, id.height, { inversionAttempts: 'attemptBoth' });
              return r ? r.data : null;
            } catch (e) { return null; }
          };
          var out = tryCanvas(c);
          if (!out && Math.min(cw, ch) < 700) {           // 小图二维码放大再试一次
            var u = document.createElement('canvas'); u.width = cw * 2; u.height = ch * 2;
            var ux = u.getContext('2d'); ux.imageSmoothingEnabled = true; ux.drawImage(c, 0, 0, cw * 2, ch * 2);
            out = tryCanvas(u);
          }
          res(out);
        } catch (e) { res(null); }
      };
      img.onerror = function () { res(null); };
      img.src = dataUrl;
    });
  });
}
/* 直接对已经渲染好的 canvas 解码（PDF 每页用） */
function decodeQrCanvas(canvas) {
  if (!window.jsQR) return null;
  try {
    var id = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height);
    var r = window.jsQR(id.data, id.width, id.height, { inversionAttempts: 'attemptBoth' });
    return r ? r.data : null;
  } catch (e) { return null; }
}

/* 解析国税局发票二维码明文串 -> 结构化字段 */
function parseInvoiceQr(str) {
  if (!str) return null;
  var s = String(str).trim();

  /* 情形 A：网址型二维码（区块链/交付码），含 bill_num / total_amount */
  var urlNum = s.match(/bill_num[=:](\d+)/i);
  var urlAmt = s.match(/total_amount[=:](\d+(?:\.\d+)?)/i);
  if (/^https?:\/\//i.test(s) && (urlNum || urlAmt)) {
    return { valid: true, type: '电子发票', code: '', number: urlNum ? urlNum[1] : '',
      date: '', total: urlAmt ? num(urlAmt[1]) : 0, check: '', raw: s };
  }

  /* 情形 B：标准逗号分隔 01,类型,代码,号码,金额,日期,校验码[,CRC] */
  var f = s.split(',').map(function (x) { return x.trim(); });
  if (f.length < 5) return null;
  function isCode(v) { return /^\d{10,12}$/.test(v); }
  function isNum(v) { return /^\d{8,20}$/.test(v); }
  function isDate(v) { return /^\d{8}$/.test(v) && +v.slice(0, 4) >= 2000 && +v.slice(4, 6) <= 12 && +v.slice(6, 8) <= 31; }
  function isMoney(v) { return /^\d+(?:\.\d+)?$/.test(v) && num(v) > 0; }
  function fmtDate(v) { return v.replace(/^(\d{4})(\d{2})(\d{2})$/, '$1-$2-$3'); }

  var code = '', number = '', total = 0, date = '', check = '';
  if (f[0] === '01' || f.length >= 6) {           // 按标准位取：f2代码 f3号码 f4金额 f5日期 f6校验码
    if (isCode(f[2])) code = f[2];
    if (isNum(f[3])) number = f[3];
    if (isMoney(f[4])) total = num(f[4]);
    if (isDate(f[5])) date = fmtDate(f[5]);
    if (f[6] && /^[0-9A-Za-z]{4,}$/.test(f[6])) check = f[6];
  }
  /* 启发式兜底：在标准位没取到时，到所有段里按类型匹配 */
  if (!code) for (var i = 0; i < f.length; i++) if (isCode(f[i])) { code = f[i]; break; }
  if (!number) for (var j = 0; j < f.length; j++) if (isNum(f[j]) && f[j] !== code) { number = f[j]; break; }
  if (!total) for (var k = 0; k < f.length; k++) if (isMoney(f[k])) { total = num(f[k]); break; }
  if (!date) for (var m = 0; m < f.length; m++) if (isDate(f[m])) { date = fmtDate(f[m]); break; }

  var typeName = QR_TYPE[f[1]] || (code ? '增值税发票' : '');
  if (!code && !number && !total && !date) return null;
  return { valid: true, type: typeName, code: code, number: number, date: date, total: total, check: check, raw: s };
}

/* OCR 结果与二维码结果合并：二维码对代码/号码/日期/金额具有权威性 */
function mergeInvoice(ocr, qr) {
  var d = {
    type: (qr && qr.type) || ocr.type || '',
    code: (qr && qr.code) || ocr.code || '',
    number: (qr && qr.number) || ocr.number || '',
    date: (qr && qr.date) || ocr.date || '',
    seller: (ocr && ocr.seller) || '',
    item: (ocr && ocr.item) || '',
    total: 0, totalSource: null
  };
  if (qr && qr.total > 0) { d.total = qr.total; d.totalSource = 'qr'; }
  else if (ocr && ocr.total) { d.total = num(ocr.total); d.totalSource = ocr.totalGuessed ? 'guess' : 'ocr'; }
  d.hasCore = !!(d.code || d.number || d.date || d.total);
  return d;
}
function applyMerged(row, d) {
  /* 摘要 = 发票出具方公司全称 + 发票内容（如「××公司 服务费」） */
  var sum = (d.seller || '').trim();
  if (d.item && d.item.trim() && d.item.trim() !== sum) sum += (sum ? ' ' : '') + d.item.trim();
  if (!sum) sum = d.type || '发票';
  row.summary = sum;
  row.amount = num(d.total);
  row.digits = amountToDigits(row.amount);
  /* 科目：专票 / 普票（按识别结果填写，可在表格手改或清空） */
  row.subject = d.vat || '';
  row.count = 1;
  row.qrOk = (d.totalSource === 'qr');
}

/* ==================================================================
 * 六、各格式文本提取
 * ================================================================== */
function preprocessForOcr(dataUrl) {
  return new Promise(function (res) {
    var img = new Image();
    img.onload = function () {
      try {
        var w = img.naturalWidth, h = img.naturalHeight;
        if (!w || !h) return res(dataUrl);
        var maxd = 1700, scale = Math.min(1, maxd / Math.max(w, h));
        var scaleUp = 1;
        if (Math.max(w, h) < 700) scaleUp = Math.min(2, 700 / Math.max(w, h));
        var tw = Math.round(w * scale * scaleUp), th = Math.round(h * scale * scaleUp);
        var c = document.createElement('canvas');
        c.width = tw; c.height = th;
        var ctx = c.getContext('2d');
        ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, tw, th);
        ctx.drawImage(img, 0, 0, tw, th);
        var d = ctx.getImageData(0, 0, tw, th), a = d.data;
        for (var i = 0; i < a.length; i += 4) {
          var g = a[i] * 0.299 + a[i + 1] * 0.587 + a[i + 2] * 0.114;
          g = g < 120 ? Math.max(0, g * 0.72) : Math.min(255, (g - 120) * 1.28 + 150);
          a[i] = a[i + 1] = a[i + 2] = g;
        }
        ctx.putImageData(d, 0, 0);
        res(c.toDataURL('image/png'));
      } catch (e) { res(dataUrl); }
    };
    img.onerror = function () { res(dataUrl); };
    img.src = dataUrl;
  });
}

async function ocrImage(dataUrl, rec) {
  var ok = await loadLib('tesseract');
  if (!ok) throw new Error('图片识别组件不可用（网络受限），请手动填写金额');
  var src = await preprocessForOcr(dataUrl);
  var opts = {
    logger: function (m) {
      if (!rec) return;
      var map = {
        'loading tesseract core': '准备识别核心',
        'initializing tesseract': '初始化',
        'loading language traineddata': '加载中文包(约6MB,仅首次)',
        'initializing api': '初始化接口',
        'recognizing text': '正在识别'
      };
      var label = map[m.status] || m.status;
      rec.status = 'ing';
      rec.error = label + (m.progress ? ' ' + Math.round(m.progress * 100) + '%' : '');
      renderListThrottled();
    }
  };
  if (ocrBase) { opts.workerPath = ocrBase.workerPath; opts.corePath = ocrBase.corePath; opts.langPath = ocrBase.langPath; }
  var out = await window.Tesseract.recognize(src, 'chi_sim', opts);
  return (out && out.data && out.data.text) || '';
}

/* 抽取 PDF 每一页的文字层 + 坐标 + 二维码，返回 pages 数组（不再按页当成一张发票） */
async function processPdf(file) {
  var ok = await loadLib('pdfjs');
  if (!ok) throw new Error('PDF 解析组件不可用（网络受限），请手动填写');
  await loadLib('jsqr');                       // 确保二维码解码库就绪
  var pl = window.pdfjsLib;
  if (pdfWorkerUrl) pl.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  var buf = await readAsArrayBuffer(file);
  var pdf = await pl.getDocument({ data: buf, isEvalSupported: false }).promise;
  var pages = [];
  for (var p = 1; p <= pdf.numPages; p++) {
    var page = await pdf.getPage(p);
    var content = await page.getTextContent();
    var str = content.items.map(function (it) { return it.str; }).join(' ');
    var items = null, qr = null;
    if (flat(str).length < 30) {
      /* 文字太少（扫描件/图片型）：渲染后扫二维码 + OCR */
      var viewport = page.getViewport({ scale: 2 });
      var canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      var ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: viewport }).promise;
      try { qr = decodeQrCanvas(canvas); } catch (e) {}
      try { str = await ocrImage(canvas.toDataURL('image/png'), null); } catch (e) { str = ''; }
    } else {
      items = content.items;                      // 保留文字层坐标，供版面感知解析（定位销售方/货物名）
      try {
        var vp = page.getViewport({ scale: 1.5 });
        var cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
        var cctx = cv.getContext('2d'); cctx.fillStyle = '#fff'; cctx.fillRect(0, 0, cv.width, cv.height);
        await page.render({ canvasContext: cctx, viewport: vp }).promise;
        qr = decodeQrCanvas(cv);
      } catch (e) {}
    }
    pages.push({ text: str, items: items, qr: qr });
  }
  return pages;
}

/* 精确识别一个 PDF 里的发票数量：按「发票号码 / 票据号码」在全文的位置切分。
   - 命中 0~1 个号码：视为单张发票（长发票跨多页也合并为一块）。
   - 命中 ≥2 个号码：每张发票一个块（块文本 = 该号码到下个号码之间），块坐标只在「该页仅此一张」时用于版面感知。 */
function splitIntoInvoices(pages) {
  var full = '', marks = [], pageMarks = {};
  pages.forEach(function (pg, pi) {
    var t = pg.text || '';
    var start = full.length;
    full += t + '\n';
    pageMarks[pi] = 0;
    var re = /发票号码[:：]?\s*([0-9]{8,20})|票据号码[:：]?\s*([0-9]{8,20})/g, m;
    while ((m = re.exec(t)) !== null) {
      marks.push({ idx: start + m.index, page: pi, number: m[1] || m[2] });
      pageMarks[pi]++;
    }
  });
  if (marks.length <= 1) {
    var txt = pages.map(function (pg) { return pg.text || ''; }).join('\n');
    var its = [];
    pages.forEach(function (pg) { if (pg.items) its = its.concat(pg.items); });
    var qr = null;
    pages.forEach(function (pg) { if (pg.qr && !qr) qr = pg.qr; });
    return [{ text: txt, items: its, qr: qr, number: marks.length ? marks[0].number : '', useLayout: true }];
  }
  var blocks = [];
  for (var k = 0; k < marks.length; k++) {
    var a = marks[k].idx;
    var b = (k + 1 < marks.length) ? marks[k + 1].idx : full.length;
    var seg = full.slice(a, b);
    var pg = pages[marks[k].page];
    blocks.push({
      text: seg,
      items: pg ? pg.items : null,
      qr: pg ? pg.qr : null,
      number: marks[k].number,
      useLayout: pageMarks[marks[k].page] === 1   // 该页仅一张发票时坐标才可信
    });
  }
  return blocks;
}

async function processExcel(file) {
  var ok = await loadLib('xlsx');
  if (!ok) throw new Error('Excel 解析组件不可用（网络受限），请手动填写');
  var X = window.XLSX;
  var buf = await readAsArrayBuffer(file);
  var wb = X.read(buf, { type: 'array' });
  var texts = [];
  wb.SheetNames.forEach(function (sn) {
    var rows = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: '', blankrows: false });
    rows.forEach(function (row) {
      var line = row.map(function (c) { return String(c == null ? '' : c); }).join(' ');
      if (flat(line)) texts.push(line);
    });
  });
  return texts.join('\n');
}

async function processWord(file) {
  var ok = await loadLib('mammoth');
  if (!ok) throw new Error('Word 解析组件不可用（网络受限），请手动填写');
  var buf = await readAsArrayBuffer(file);
  var res = await window.mammoth.extractRawText({ arrayBuffer: buf });
  return res.value || '';
}

/* ==================================================================
 * 七、上传与识别
 * ================================================================== */
function takeBlankRow() {
  for (var i = 0; i < state.rows.length; i++) {
    var r = state.rows[i];
    if (!r.invId && !r.summary && !r.subject && !num(r.count) && !num(r.amount)) return r;
  }
  var nr = blankRow();
  state.rows.push(nr);
  return nr;
}
/* （applyParsed 已由 mergeInvoice + applyMerged 取代，见五·a） */

/* 本次会话内被用户选择「重新报销」放行的发票文件指纹集合 */
var allowedForce = new Set();

/* 弹出重复上传提醒窗，返回 Promise<true=重新报销 / false=谢谢提醒> */
function showDuplicateModal(msg) {
  return new Promise(function (resolve) {
    var m = $('#dupModal');
    if (!m) { resolve(false); return; }
    $('#dupMsg').textContent = msg;
    m.classList.remove('hidden');
    var bRe = $('#dupReReimb'), bTh = $('#dupThanks');
    function done(choice) {
      m.classList.add('hidden');
      bRe.removeEventListener('click', onRe);
      bTh.removeEventListener('click', onTh);
      resolve(choice);
    }
    function onRe() { done(true); }
    function onTh() { done(false); }
    bRe.addEventListener('click', onRe);
    bTh.addEventListener('click', onTh);
  });
}

async function handleFiles(fileList) {
  var files = Array.prototype.slice.call(fileList);
  if (!files.length) { toast('没有选到文件'); return; }
  var accept = ['pdf', 'jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif', 'doc', 'docx', 'xls', 'xlsx', 'csv'];
  var lg = loadReimb();                       // 仅含「已导出」台账
  var jobs = [], blocked = [], skipped = [], seen = {};
  for (var fi = 0; fi < files.length; fi++) {
    var f = files[fi];
    var ext = extOf(f.name);
    if (accept.indexOf(ext) < 0) { skipped.push(f.name); continue; }
    var hash = await fileHash(f);
    if (seen[hash]) { skipped.push('「' + f.name + '」本批已处理'); continue; }
    seen[hash] = true;
    /* 仅与「已导出」台账比对：未导出过（含误删后重新上传）不再拦截 */
    if (lg.hash[hash]) {
      if (allowedForce.has(hash)) {
        /* 本次会话已选过「重新报销」此文件，直接放行 */
      } else {
        var dupMsg = '「' + f.name + '」已于 ' + fmtLedgerDate(lg.hash[hash].when) + ' 导出过（文件名『' + lg.hash[hash].fileName + '』）。是否仍要重新报销？';
        var goAhead = await showDuplicateModal(dupMsg);
        if (!goAhead) { blocked.push('「' + f.name + '」已拦截重复报销'); continue; }
        allowedForce.add(hash);
      }
    }
    jobs.push({ file: f, hash: hash, ext: ext });
  }
  renderAll();
  if (skipped.length) toast('忽略：' + skipped.join('、'));
  if (blocked.length) toast(blocked.join('；'));
  if (!jobs.length) return;
  toast('已接收 ' + jobs.length + ' 个文件，开始识别…（首张需先准备 OCR 组件）');
  for (var i = 0; i < jobs.length; i++) await processFile(jobs[i]);
  renderAll();
  var okN = state.invoices.filter(function (r) { return r.status === 'ok'; }).length;
  toast('识别完成：成功 ' + okN + ' 张，其余 ' + (state.invoices.length - okN) + ' 张可在表格里手工补填');
}

/* 一个上传文件 -> 按内容拆成 N 张发票（PDF 精确切分；图片/Word/Excel 通常为 1 张），
   每张发票各建一个 rec + 一行，识别后自动补位填入报销单。 */
async function processFile(job) {
  var f = job.file, ext = job.ext;
  var blocks;
  try {
    if (ext === 'pdf') {
      var pages = await processPdf(f);
      blocks = splitIntoInvoices(pages);
    } else if (['jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif'].indexOf(ext) >= 0) {
      var dataUrl = await readAsDataURL(f);
      var iqr = decodeQr(dataUrl);
      var itext = await ocrImage(dataUrl, null);
      blocks = [{ text: itext, items: null, qr: iqr, number: '', useLayout: false }];
    } else if (['xls', 'xlsx', 'csv'].indexOf(ext) >= 0) {
      blocks = [{ text: await processExcel(f), items: null, qr: null, number: '', useLayout: false }];
    } else if (['doc', 'docx'].indexOf(ext) >= 0) {
      blocks = [{ text: await processWord(f), items: null, qr: null, number: '', useLayout: false }];
    } else {
      blocks = [];
    }
  } catch (e) {
    var failRec = { id: uid(), fileName: f.name, ext: ext, thumb: '', status: 'todo', error: (e && e.message) ? e.message : '识别失败', rawText: '', progress: 0, qrOk: false, file: f, hash: job.hash, number: '' };
    if (['jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif'].indexOf(ext) >= 0) { try { failRec.thumb = await readAsDataURL(f); } catch (e2) {} }
    state.invoices.push(failRec);
    var frow = takeBlankRow(); frow.invId = failRec.id; frow.summary = baseName(f.name); frow.count = 1;
    renderAll();
    return;
  }
  if (!blocks.length) return;
  var n = blocks.length, recs = [];
  for (var k = 0; k < n; k++) {
    var b = blocks[k];
    var rec = {
      id: uid(),
      fileName: n > 1 ? (baseName(f.name) + '（' + (k + 1) + '/' + n + '）') : f.name,
      ext: ext, thumb: '', status: 'ing', error: '排队中', rawText: b.text || '', progress: 0, qrOk: !!(b.qr && b.qr.valid), file: f, hash: job.hash, number: b.number || ''
    };
    if (['jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif'].indexOf(ext) >= 0) { try { rec.thumb = await readAsDataURL(f); } catch (e3) {} }
    state.invoices.push(rec);
    var row = takeBlankRow(); row.invId = rec.id; row.summary = baseName(rec.fileName); row.count = 1;
    recs.push({ rec: rec, row: row, block: b });
  }
  renderAll();
  for (var j = 0; j < recs.length; j++) await processOne(recs[j]);
}

async function processOne(job) {
  var rec = job.rec, ext = job.ext;
  rec.status = 'ing'; rec.error = '识别中…';
  renderListThrottled();
  try {
    var b = job.block;
    var text = b.text || '', qr = b.qr, pdfItems = b.items;
    rec.rawText = text;
    if (qr) rec.qrOk = true;
    var parsed = parseInvoice(text || '');
    /* 电子发票文字层：仅当该发票独占一页（坐标可靠）时用坐标定位销售方/货物名；
       同页多张时坐标会串，依赖文字层解析即可 */
    if (pdfItems && pdfItems.length && b.useLayout) {
      var lay = parsePdfLayout(pdfItems);
      /* 版面感知只补文字层取不到的字段：文字层已取到销售方/货物名就不覆盖，
         避免坐标误把「开户行账号」当销售方、或把货物名截断。 */
      if (!parsed.seller && lay.seller) parsed.seller = lay.seller;
      if (!parsed.item && lay.item) parsed.item = lay.item;
    }
    var data = mergeInvoice(parsed, qr);
    data.vat = detectVat(text, qr);           // 科目：专票/普票（机器优先）
    rec.number = data.number || rec.number || '';
    /* 重复报销拦截：只比对「已导出」台账的票号；同会话列表里重复同号仅提示，不强制。
       台账只在导出时由 markExported 写入，上传时不再写 —— 误删后重传不再误报。 */
    if (rec.number && (data.hasCore || data.seller || data.item)) {
      var lg2 = loadReimb();
      var prev = lg2.num[rec.number];
      var sameInState = state.invoices.some(function (x) { return x.id !== rec.id && x.number === rec.number; });
      if (prev && prev.hash !== rec.hash) {
        if (allowedForce.has(rec.hash)) {
          /* 用户在 handleFiles 已选「重新报销」，本次直接放行 */
        } else {
          var re2 = await showDuplicateModal('发票（号码 ' + rec.number + '）已于 ' + fmtLedgerDate(prev.when) + ' 导出过（当时文件名『' + prev.fileName + '』）。是否仍要重新报销？');
          if (re2) {
            allowedForce.add(rec.hash);
            lg2.num[rec.number] = { when: isoToday(), fileName: rec.fileName, hash: rec.hash };
            saveReimb(lg2);
          } else {
            blockDuplicate(rec, job.row, prev);
            return;
          }
        }
      } else if (sameInState) {
        toast('「' + rec.fileName + '」的发票号码与列表里另一张重复，请核对');
      }
    }
    if (data.hasCore || data.seller || data.item) {
      applyMerged(job.row, data);
      if (data.date) state.date = normDate(data.date);   // 报销日期：按发票开票日期填写
      if (num(job.row.amount) > MAX_GRID) {
        toast('「' + rec.fileName + '」金额超过 999,999.99，分格只显示到十万位，合计仍按真实金额计算');
      }
      rec.status = 'ok';
      rec.error = data.totalSource === 'qr'
        ? '二维码已识别（票号/金额来自二维码）'
        : data.totalSource === 'guess' ? '金额为推测值，请核对' : '已识别';
      if (data.totalSource === 'guess') rec.status = 'todo';
      if (qr && qr.valid) toast('二维码识别成功：' + (qr.number ? '票号' + qr.number + ' ' : '') + '金额 ¥' + capital2(qr.total));
    } else {
      rec.status = 'todo';
      rec.error = rec.rawText ? '未识别到发票字段，请手工补填' : '未提取到文字，请手工补填';
    }
  } catch (e) {
    rec.status = 'todo';
    rec.error = (e && e.message) ? e.message : '识别失败，请手工补填';
  }
  renderList();
  renderSheet();
  updateTotals();
}

var lastListRender = 0, listTimer = null;
function renderListThrottled() {
  var now = Date.now();
  if (now - lastListRender > 320) { lastListRender = now; renderList(); return; }
  if (listTimer) return;
  listTimer = setTimeout(function () { listTimer = null; lastListRender = Date.now(); renderList(); }, 340);
}

/* ==================================================================
 * 八、渲染
 * ================================================================== */
/* 删除发票（与 ✕ 删行一致）：移除发票记录并 splice 对应明细行，自动补位、清除多余空白报销单 */
function removeInvoice(id) {
  state.invoices = state.invoices.filter(function (x) { return x.id !== id; });
  for (var di = state.rows.length - 1; di >= 0; di--) {
    if (state.rows[di].invId === id) state.rows.splice(di, 1);
  }
  if (!state.rows.length) { for (var bi = 0; bi < 5; bi++) state.rows.push(blankRow()); }
  renderAll();
}

/* 渲染「已上传发票」详情弹窗：完整文件名 + 状态 + 删除；支持按文件名中文子串搜索 */
function renderInvoiceDetail() {
  var ul = $('#invDetailList'); if (!ul) return;
  var q = ($('#invSearch') && $('#invSearch').value) ? $('#invSearch').value.trim().toLowerCase() : '';
  var cnt = $('#invDetailCount'); if (cnt) cnt.textContent = state.invoices.length;
  ul.innerHTML = '';
  var vis = 0;
  state.invoices.forEach(function (rec) {
    if (q && (rec.fileName || '').toLowerCase().indexOf(q) < 0) return;
    vis++;
    var li = document.createElement('li');
    li.className = 'inv-detail-item';
    var tag = rec.status === 'ok'
      ? (rec.qrOk ? '<span class="tag ok">二维码已识别</span>' : '<span class="tag ok">已识别</span>')
      : rec.status === 'ing' ? '<span class="tag ing">识别中</span>'
      : '<span class="tag bad">待补填</span>';
    li.innerHTML =
      '<div class="inv-d-name" title="' + esc(rec.fileName) + '">' + esc(rec.fileName) + '</div>' +
      '<div class="inv-d-meta">' + tag + '</div>' +
      '<button class="icon-btn del" type="button" data-del="' + rec.id + '" title="删除">删除</button>';
    ul.appendChild(li);
  });
  var empty = $('#invDetailEmpty'); if (empty) empty.classList.toggle('hidden', vis !== 0);
}

function renderAll() { renderList(); renderSheet(); updateTotals(); renderInvoiceDetail(); }

function renderList() {
  var ul = $('#invoiceList');
  if (!ul) return;
  var cnt = $('#invCount'); if (cnt) cnt.textContent = state.invoices.length;
  var clr = $('#clearAll'); if (clr) clr.hidden = state.invoices.length === 0;
  ul.innerHTML = '';
  state.invoices.forEach(function (rec) {
    var li = document.createElement('li');
    li.className = 'inv-item';
    var tag = rec.status === 'ok'
      ? (rec.qrOk ? '<span class="tag ok">二维码已识别</span>' : '<span class="tag ok">已识别</span>')
      : rec.status === 'ing'
        ? '<span class="tag ing">识别中</span>'
        : '<span class="tag err">待补填</span>';
    var thumb = rec.thumb
      ? '<div class="thumb"><img src="' + rec.thumb + '" alt=""></div>'
      : '<div class="thumb">' + esc((rec.ext || 'FILE').toUpperCase()) + '</div>';
    var sub = tag + (rec.error ? ' <span>' + esc(rec.error) + '</span>' : '');
    li.innerHTML = thumb +
      '<div class="meta">' +
      '<div class="name" title="' + esc(rec.fileName) + '">' + esc(rec.fileName) + '</div>' +
      '<div class="sub">' + sub + '</div>' +
      '</div>' +
      '<div class="ops">' +
      '<button class="icon-btn" type="button" data-view="' + rec.id + '">原文</button>' +
      '<button class="icon-btn" type="button" data-retry="' + rec.id + '">重识别</button>' +
      '<button class="icon-btn del" type="button" data-del="' + rec.id + '">删除</button>' +
      '</div>';
    ul.appendChild(li);
  });
}

/* 列宽取自原 xlsx（B12 C12 D18.6 E~L各3 M9.13 N8.13，合计 83.86） */
var COL_PCT = [14.31, 14.31, 22.18, 3.577, 3.577, 3.577, 3.577, 3.577, 3.577, 3.577, 3.577, 10.89, 9.70];

var FORM_ROWS = 5;   // 每张报销单 5 行明细；超出自动生成第 2、3…张
/* 把 flat 的 state.rows 切成「每张 5 行」的若干报销单；最后一张不足 5 行也补齐到 5 行 */
function formsFromRows() {
  var per = FORM_ROWS;
  var rows = state.rows.slice();
  var n = Math.max(1, Math.ceil(rows.length / per));
  var forms = [];
  for (var f = 0; f < n; f++) {
    var chunk = rows.slice(f * per, f * per + per);
    while (chunk.length < per) chunk.push(blankRow());
    forms.push(chunk);
  }
  return forms;
}

/* 生成「单张报销单」的完整 <table> 内部 HTML。
   screen=true 含删除按钮与可编辑输入框；screen=false 供导出用（仍输出 input，由 cleanFormInputs 替换为纯文字）。 */
function buildFormHtml(rows, opt) {
  opt = opt || {};
  var base = opt.base || 0;           // 该张明细在 state.rows 中的全局起始下标
  var k = (opt.k != null) ? opt.k : 0;
  var screen = !!opt.screen;
  var html = '';

  html += '<colgroup>';
  COL_PCT.forEach(function (w) { html += '<col style="width:' + w + '%">'; });
  html += '</colgroup>';

  /* 行1 标题 */
  html += '<tr><td class="bx-title" colspan="13">苏州沛斯仁光电科技有限公司费用报销单</td></tr>';

  /* 行2 部门 | 报销日期（每张都渲染，绑定同一份共享值） */
  html += '<tr>' +
    '<td class="bx-h" colspan="2">部门：<input class="bx-in bx-line f-dept" type="text" data-k="dept" value="' + esc(state.dept) + '"></td>' +
    '<td class="bx-h bx-h-right" colspan="11">报销日期：<input class="bx-in f-date" type="text" data-k="date" value="' + esc(state.date) + '" placeholder="按发票开票日期"></td>' +
    '</tr>';

  /* 行3-4 表头 */
  html += '<tr>' +
    '<td class="bx-th" colspan="3" rowspan="2">摘要</td>' +
    '<td class="bx-th" colspan="8">金额</td>' +
    '<td class="bx-th bx-subject" rowspan="2">科目</td>' +
    '<td class="bx-th" rowspan="2">单据<br>张数</td>' +
    '</tr><tr>';
  DIGIT_LABELS.forEach(function (l) { html += '<td class="bx-ths">' + l + '</td>'; });
  html += '</tr>';

  /* 行5.. 明细（每一行都是可直接输入的实体行，全局下标 = base + 局部下标） */
  rows.forEach(function (row, li) {
    var gi = base + li;
    html += '<tr class="bx-row" data-r="' + gi + '">' +
      '<td class="bx-summary" colspan="3">' +
      (screen ? '<button class="bx-del screen-only" type="button" data-r="' + gi + '" title="删除本行">✕</button>' : '') +
      '<input class="bx-in" type="text" data-r="' + gi + '" data-k="summary" value="' + esc(row.summary) + '">' +
      '</td>';
    var vd = visibleDigits(row.digits);
    for (var d = 0; d < 8; d++) {
      html += '<td class="bx-digit"><input class="bx-in" type="text" maxlength="1" inputmode="numeric" data-r="' + gi + '" data-d="' + d + '" value="' + esc(vd[d] || '') + '"></td>';
    }
    html += '<td class="bx-subject">' + '<input class="bx-in" type="text" data-r="' + gi + '" data-k="subject" value="' + esc(row.subject) + '">' + '</td>';
    html += '<td><input class="bx-in bx-incount" type="text" maxlength="3" inputmode="numeric" data-r="' + gi + '" data-k="count" value="' + esc(row.count) + '"></td>';
    html += '</tr>';
  });

  /* 合计（大写）行：仅合计本张明细 */
  var ftot = 0; rows.forEach(function (r) { ftot += num(r.amount); });
  var fcnt = 0; rows.forEach(function (r) { fcnt += num(r.count); });
  var td = amountToDigits(ftot);
  var vtd = visibleDigits(td);
  html += '<tr class="bx-total">' +
    '<td class="bx-tl" colspan="3">合计（大写）：<b class="t-capital" data-form="' + k + '">' + toCapital(ftot) + '</b></td>';
  vtd.forEach(function (v, i) {
    html += '<td class="bx-digit bx-tot"><span class="t-totd" data-form="' + k + '" data-i="' + i + '">' + (v || '') + '</span></td>';
  });
  html += '<td class="bx-tc" colspan="2">单据 <b class="t-tcount" data-form="' + k + '">' + Math.round(fcnt) + '</b></td></tr>';

  /* 签字栏（严格按原文件列位：部门主管合并 6 列、复核合并 3 列）；每张都渲染，绑定同一份共享值 */
  function sgColspan(label, key, cs) {
    return '<td colspan="' + cs + '"><div class="sg-cell"><span class="sg-l">' + label + '</span>' +
      '<input class="bx-in sg-in" type="text" data-s="' + key + '" value="' + esc(state.sign[key]) + '"></div></td>';
  }
  html += '<tr class="bx-sign">' +
    sgColspan('财会主管', 'caik', 1) +
    sgColspan('记账', 'jizhang', 1) +
    sgColspan('经办人', 'jingban', 1) +
    sgColspan('部门主管', 'bumen', 6) +
    sgColspan('复核', 'fuhe', 3) +
    sgColspan('报销人', 'baoxiao', 1) +
    '</tr>';
  return html;
}

function renderSheet() {
  var formsEl = $('#forms');
  if (!formsEl) return;
  var forms = formsFromRows();
  var html = '';
  forms.forEach(function (rows, k) {
    html += '<div class="form-block">' +
      '<div class="form-cap screen-only">报销单 ' + (k + 1) + (forms.length > 1 ? ' / ' + forms.length : '') + '</div>' +
      '<table class="bx-table" data-form="' + k + '">' +
      buildFormHtml(rows, { screen: true, base: k * FORM_ROWS, k: k }) +
      '</table></div>';
  });
  formsEl.innerHTML = html;
  /* 每张表内：明细行行高统一（空白行撑到与已填行同高） */
  Array.prototype.forEach.call(formsEl.querySelectorAll('.bx-table'), unifyDetailRows);
  var et = $('#emptyTip'); if (et) et.classList.toggle('hidden', hasData());
}

/* 明细行行高统一：以「已填入发票内容的整行」高度为标准，空白行也撑到同样高，
   避免填了内容的行比空行高（屏上、导出 PDF/Excel/Word/打印 都统一） */
function unifyDetailRows(table) {
  if (!table) return;
  var rows = Array.prototype.slice.call(table.querySelectorAll('tr.bx-row'));
  if (rows.length < 2) return;
  rows.forEach(function (r) { r.style.height = ''; });
  var max = 0;
  rows.forEach(function (r) {
    var h = r.getBoundingClientRect().height || r.offsetHeight || 0;
    if (h > max) max = h;
  });
  if (max > 0) rows.forEach(function (r) { r.style.height = max + 'px'; });
}

function updateTotals() {
  var forms = formsFromRows();
  forms.forEach(function (rows, k) {
    var tot = 0; rows.forEach(function (r) { tot += num(r.amount); });
    var cap = document.querySelector('.t-capital[data-form="' + k + '"]');
    if (cap) cap.textContent = toCapital(tot);
    var ds = visibleDigits(amountToDigits(tot));
    for (var i = 0; i < 8; i++) {
      var el = document.querySelector('.t-totd[data-form="' + k + '"][data-i="' + i + '"]');
      if (el) el.textContent = ds[i] || '';
    }
    var fc = 0; rows.forEach(function (r) { fc += num(r.count); });
    var tc = document.querySelector('.t-tcount[data-form="' + k + '"]');
    if (tc) tc.textContent = Math.round(fc);
  });
  var et = $('#emptyTip'); if (et) et.classList.toggle('hidden', hasData());
}

/* ==================================================================
 * 九、编辑交互
 * ================================================================== */
function bindSheet() {
  var t = $('#forms');
  if (!t) return;

  t.addEventListener('input', function (e) {
    var el = e.target;
    if (!el || !el.dataset) return;
    if (el.dataset.k === 'dept') { state.dept = el.value; return; }
    if (el.dataset.k === 'date') { state.date = el.value; return; }
    if (el.dataset.d !== undefined) {
      var r = state.rows[+el.dataset.r];
      if (!r) return;
      el.value = sanitizeDigit(el.value);
      r.digits[+el.dataset.d] = el.value;
      r.amount = digitsToAmount(r.digits);
      updateTotals();
    } else if (el.dataset.s) {
      state.sign[el.dataset.s] = el.value;
    } else if (el.dataset.k) {
      var r2 = state.rows[+el.dataset.r];
      if (!r2) return;
      var k = el.dataset.k;
      if (k === 'count') { el.value = String(el.value).replace(/[^0-9]/g, ''); r2.count = el.value; updateTotals(); }
      else { r2[k] = el.value; }
      if (k === 'summary' || k === 'subject') updateTotals();
    }
  });

  t.addEventListener('click', function (e) {
    var el = e.target;
    if (el && el.classList && el.classList.contains('bx-del')) {
      var i = +el.dataset.r;
      if (!isNaN(i)) deleteRow(i);
    }
  });
}

function deleteRow(i) {
  if (state.rows.length <= 1) {
    state.rows[0] = blankRow();          // 最后一行改为清空
  } else {
    var removed = state.rows.splice(i, 1)[0];
    if (removed && removed.invId) {
      state.invoices = state.invoices.filter(function (x) { return x.id !== removed.invId; });
    }
  }
  renderAll();
}
function addRow() {
  state.rows.push(blankRow());
  renderAll();
  toast('已增加一行，共 ' + state.rows.length + ' 行明细');
}

function bindSidebar() {
  var dz = $('#dropzone'), input = $('#fileInput'), pick = $('#pickBtn');
  pick.addEventListener('click', function () { input.click(); });
  input.addEventListener('change', function () {
    var fs = input.files;
    handleFiles(fs);
    try { input.value = ''; } catch (err) {}
  });
  ['dragenter', 'dragover'].forEach(function (ev) {
    dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('drag'); });
  });
  ['dragleave', 'drop'].forEach(function (ev) {
    dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('drag'); });
  });
  dz.addEventListener('drop', function (e) {
    if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });

  $('#invoiceList').addEventListener('click', function (e) {
    var el = e.target;
    if (!el || !el.dataset) return;
    if (el.dataset.view) {
      var rec = state.invoices.filter(function (x) { return x.id === el.dataset.view; })[0];
      if (rec) alert('【' + rec.fileName + '】识别到的原文：\n\n' + (rec.rawText && flat(rec.rawText) ? rec.rawText : '（没有提取到文字，可能是模糊图片，请在表格里手工填写）'));
    } else if (el.dataset.retry) {
      retryOne(el.dataset.retry);
    } else if (el.dataset.del) {
      var id = el.dataset.del;
      var r = state.invoices.filter(function (x) { return x.id === id; })[0];
      if (r && confirm('删除这张发票？对应的明细行也会一并删除。')) removeInvoice(id);
    }
  });

  $('#invDetailList').addEventListener('click', function (e) {
    var el = e.target;
    if (!el || !el.dataset) return;
    if (el.dataset.del) {
      var id = el.dataset.del;
      var r = state.invoices.filter(function (x) { return x.id === id; })[0];
      if (r && confirm('删除这张发票？对应的明细行也会一并删除。')) removeInvoice(id);
    }
  });
  var openD = $('#openInvDetail');
  if (openD) openD.addEventListener('click', function () {
    renderInvoiceDetail();
    $('#invDetailModal').classList.remove('hidden');
    if ($('#invSearch')) $('#invSearch').value = '';
  });
  $$('[data-close-inv]').forEach(function (b) { b.addEventListener('click', function () { $('#invDetailModal').classList.add('hidden'); }); });
  var ivs = $('#invSearch');
  if (ivs) ivs.addEventListener('input', renderInvoiceDetail);

  $('#clearAll').addEventListener('click', function () {
    if (!state.invoices.length) return;
    if (confirm('清空全部已上传发票及其明细行？')) {
      state.invoices = [];
      state.rows = state.rows.filter(function (r) { return !r.invId && (r.summary || r.subject || num(r.count) || num(r.amount)); });
      if (!state.rows.length) { for (var i = 0; i < 5; i++) state.rows.push(blankRow()); }
      renderAll();
    }
  });

  $('#addRow').addEventListener('click', addRow);
  $('#clearRows').addEventListener('click', function () {
    if (confirm('清空报销单全部明细（包括金额）？')) {
      state.invoices = [];
      state.rows = [];
      for (var i = 0; i < 5; i++) state.rows.push(blankRow());
      renderAll();
    }
  });
  $('#reloadEngines').addEventListener('click', function () {
    toast('正在重新加载识别组件…');
    reloadEngines();
  });
}

async function retryOne(id) {
  var rec = state.invoices.filter(function (x) { return x.id === id; })[0];
  if (!rec) return;
  if (!rec.file) { toast('原文件已不在内存中，请重新上传'); return; }
  var row = state.rows.filter(function (r) { return r.invId === id; })[0];
  if (!row) { row = takeBlankRow(); row.invId = id; row.summary = baseName(rec.fileName); row.count = 1; }
  await processOne({ file: rec.file, rec: rec, row: row, ext: rec.ext });
  toast('「' + rec.fileName + '」已重新识别');
}

/* ==================================================================
 * 十、导出
 * ================================================================== */
/* 重复上传拦截：命中已报销台账则清掉占行，提示日期与当时文件名 */
function blockDuplicate(rec, row, entry) {
  state.invoices = state.invoices.filter(function (x) { return x.id !== rec.id; });
  if (row) { row.invId = null; row.summary = ''; row.subject = ''; row.count = ''; row.amount = 0; row.digits = emptyDigits(); }
  renderAll();
  toast('发票（号码 ' + rec.number + '）已于 ' + fmtLedgerDate(entry.when) + ' 上传过（当时文件名『' + entry.fileName + '』），已拦截重复上传');
}

/* 导出/打印时，把当前报销单里用到的发票标记「已报销」（下次上传即拦截） */
function markExported() {
  var lg = loadReimb();
  var when = isoToday();
  state.invoices.forEach(function (rec) {
    var row = state.rows.filter(function (r) { return r.invId === rec.id; })[0];
    var used = row && (row.summary || row.subject || num(row.count) || num(row.amount));
    if (!used) return;
    /* 仅在上传时未记录过才补写，保留最初「上传日期」作为拦截提示 */
    if (rec.hash && !lg.hash[rec.hash]) lg.hash[rec.hash] = { when: when, fileName: rec.fileName };
    if (rec.number && !lg.num[rec.number]) lg.num[rec.number] = { when: when, fileName: rec.fileName, hash: rec.hash };
  });
  saveReimb(lg);
}

/* 干净表格克隆（输入框->纯文字、去按钮），并按固定像素宽设置列宽，便于离线导出对齐 */
/* 从「某张报销单的明细行」构建一张干净、带固定列宽与可靠边框模型的 <table> 元素（供导出/出图） */
/* 把「最右列的单元格」加右边框、「最后一行的单元格」加下边框。
   原因：html2canvas 只可靠绘制单元格边框，不能依赖 <table> 元素的右/下边框。
   用 cellIndex 等价推导的「视觉列号」计算 rightCol（含 colspan），
   并用 ri+rowspan 计算 bottomRow，兼容合并单元格；LAST_COL=12（共 13 列 A..N）。 */
function addOuterBorders(t) {
  var LAST_COL = 12;
  var trs = Array.prototype.slice.call(t.querySelectorAll('tr'));
  var lastRow = trs.length - 1;
  var occupied = {};            // occupied[c] = 被上一行 rowspan 占用到的大于一行的行号
  trs.forEach(function (tr, ri) {
    var col = 0;
    Array.prototype.forEach.call(tr.children, function (td) {
      while (occupied[col] && occupied[col] > ri) col++;   // 跳过被 rowspan 占用的列
      if (td.tagName !== 'TD' && td.tagName !== 'TH') return;
      var cs = parseInt(td.getAttribute('colspan') || '1', 10) || 1;
      var rs = parseInt(td.getAttribute('rowspan') || '1', 10) || 1;
      var rightCol = col + cs - 1;
      var bottomRow = ri + rs - 1;
      if (rightCol === LAST_COL) td.style.borderRight = '1px solid #2b2b2b';
      if (bottomRow === lastRow) td.style.borderBottom = '1px solid #2b2b2b';
      if (rs > 1) { for (var c = col; c <= rightCol; c++) occupied[c] = ri + rs; }
      col = rightCol + 1;
    });
  });
}
function buildFormTable(rows, fixedPx) {
  var t = document.createElement('table');
  t.setAttribute('class', 'bx-table');
  t.innerHTML = buildFormHtml(rows, { screen: false, base: 0, k: 0 });
  /* 导出边框模型（html2canvas 友好）：
     html2canvas 不会可靠绘制 <table> 元素自身的边框（右/下边尤其容易被丢），
     所以最右边框 / 最下边框必须画在「单元格」上。
     做法：每个 td 画上/左边框（内部网格）；addOuterBorders 把右/下外框画在
     「最右列 / 最后一行」的单元格上。兼容 colspan/rowspan。 */
  t.style.borderCollapse = 'separate';
  t.style.borderSpacing = '0';
  t.style.borderRight = '0';
  t.style.borderBottom = '0';
  Array.prototype.forEach.call(t.querySelectorAll('td'), function (td) {
    td.style.borderTop = '1px solid #2b2b2b';
    td.style.borderLeft = '1px solid #2b2b2b';
    td.style.borderRight = '0';
    td.style.borderBottom = '0';
  });
  addOuterBorders(t);
  var cg = t.querySelector('colgroup');
  if (cg) {
    cg.innerHTML = '';
    COL_PCT.forEach(function (w) {
      var c = document.createElement('col');
      c.style.width = (w / 100 * fixedPx) + 'px';
      cg.appendChild(c);
    });
  }
  return t;
}
/* 离屏渲染某张报销单为 canvas（固定宽，保证 A4 比例稳定） */
function renderFormToCanvas(rows, fixedPx) {
  return loadLib('html2canvas').then(function (ok) {
    if (!ok) throw new Error('出图组件不可用（网络受限）');
    var t = buildFormTable(rows, fixedPx);
    if (!t) throw new Error('没有可导出的报销单');
    var holder = document.createElement('div');
    /* 不加 padding、用 content-box，确保表格刚好 fixedPx 宽，避免右侧被内边距挤出导致裁切 */
    holder.style.cssText = 'position:fixed;left:-99999px;top:0;width:' + fixedPx + 'px;background:#fff;box-sizing:content-box;';
    holder.appendChild(t);
    document.body.appendChild(holder);
    unifyDetailRows(t);   // 截图前对齐明细行高，保证 PDF 里空白行与已填行同高
    return new Promise(function (res, rej) {
      window.html2canvas(holder, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false }).then(function (c) {
        if (holder.parentNode) holder.parentNode.removeChild(holder);
        res(c);
      }, function (e) { if (holder.parentNode) holder.parentNode.removeChild(holder); rej(e); });
    });
  });
}
/* 单张报销单的内联化 HTML（用于 Excel / 打印，固定宽 + 显式边框） */
function oneFormTableHtml(rows, fixedPx) {
  var t = buildFormTable(rows, fixedPx);
  if (!t) return '';
  var holder = document.createElement('div');
  holder.style.cssText = 'position:fixed;left:-99999px;top:0;width:' + fixedPx + 'px;background:#fff;box-sizing:content-box';
  holder.appendChild(t); document.body.appendChild(holder);
  unifyDetailRows(t);
  cleanFormInputs(t);
  inlineize(t);
  t.setAttribute('style', 'border-collapse:collapse;width:' + fixedPx + 'px;table-layout:fixed;font-family:Microsoft YaHei,sans-serif;font-size:11pt');
  var html = t.outerHTML;
  if (holder.parentNode) holder.parentNode.removeChild(holder);
  return html;
}

/* 把表格里的输入框替换为纯文字 span（去掉操作按钮），供导出 */
function cleanFormInputs(t) {
  Array.prototype.forEach.call(t.querySelectorAll('.screen-only'), function (n) {
    if (n.parentNode) n.parentNode.removeChild(n);
  });
  Array.prototype.forEach.call(t.querySelectorAll('input,select,textarea'), function (el) {
    var sp = document.createElement('span');
    sp.className = (el.className || '') + ' bx-val';
    sp.textContent = el.value || '';
    if (el.classList && el.classList.contains('bx-line')) sp.className += ' bx-line-val';
    if (el.parentNode) el.parentNode.replaceChild(sp, el);
  });
  return t;
}
/* 给每个格子写内联样式，保证 Word/Excel 里也有框线和居中 */
function inlineize(t) {
  Array.prototype.forEach.call(t.querySelectorAll('td'), function (td) {
    var cls = ' ' + (td.className || '') + ' ';
    var st = 'border:1px solid #000000;padding:4px 3px;vertical-align:middle;';
    if (/ bx-title /.test(cls)) st += 'font-size:18pt;font-weight:bold;text-align:center;padding:8px 0;letter-spacing:0;white-space:nowrap;';
    else if (/ bx-h /.test(cls)) st += 'text-align:left;font-weight:bold;';
    else if (/ bx-tl /.test(cls)) st += 'text-align:left;font-weight:bold;';
    else if (/ bx-summary /.test(cls)) st += 'text-align:left;';
    else st += 'text-align:center;';
    if (/ bx-th | bx-ths | bx-total /.test(cls)) st += 'font-weight:bold;';
    td.setAttribute('style', st);
    td.removeAttribute('class');
  });
  Array.prototype.forEach.call(t.querySelectorAll('span.bx-val'), function (sp) {
    sp.setAttribute('style', /bx-line-val/.test(sp.className) ? 'border-bottom:1px solid #000;padding:0 40px;' : '');
    sp.removeAttribute('class');
  });
  t.removeAttribute('class');
  t.removeAttribute('id');
  t.setAttribute('border', '1');
  t.setAttribute('cellspacing', '0');
  t.setAttribute('cellpadding', '0');
  t.setAttribute('style', 'border-collapse:collapse;width:100%');
  return t;
}
function officeHtml(rows) {
  var t = buildFormTable(rows, 720);
  if (!t) return '';
  var holder = document.createElement('div');
  holder.style.cssText = 'position:fixed;left:-99999px;top:0;width:720px;background:#fff;box-sizing:content-box';
  holder.appendChild(t); document.body.appendChild(holder);
  unifyDetailRows(t);
  cleanFormInputs(t);
  inlineize(t);
  var inner = t.innerHTML;
  if (holder.parentNode) holder.parentNode.removeChild(holder);
  return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="utf-8"><title>费用报销单</title></head><body>' +
    '<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;font-family:Microsoft YaHei,sans-serif;font-size:11pt">' +
    inner + '</table></body></html>';
}
/* 用 xlsx-js-style 把「某张报销单」写成一个 worksheet（含列宽/行高/合并/边框/字体/对齐） */
function buildSheetForForm(rows, label) {
  var X = window.XLSX_STYLE;
  if (!X) return null;
  var ws = {};
  var bd = { top: { style: 'thin', color: { rgb: 'FF000000' } }, bottom: { style: 'thin', color: { rgb: 'FF000000' } }, left: { style: 'thin', color: { rgb: 'FF000000' } }, right: { style: 'thin', color: { rgb: 'FF000000' } } };
  var base = { name: '微软雅黑', sz: 10 };
  function set(r, c, val, opt) {
    opt = opt || {};
    var ref = X.utils.encode_cell({ r: r, c: c });
    var st = {
      border: bd,
      font: Object.assign({}, base, opt.font || {}),
      alignment: Object.assign({ vertical: 'center', horizontal: opt.align || 'center', wrapText: !!opt.wrap }, opt.align2 || {})
    };
    ws[ref] = { v: val, t: (typeof val === 'number') ? 'n' : 's', s: st };
  }
  /* 行1 标题 */
  set(0, 0, '苏州沛斯仁光电科技有限公司费用报销单', { align: 'center', font: { name: '微软雅黑', sz: 16, bold: true } });
  /* 行2 部门 | 报销日期 */
  set(1, 0, '部门：' + (state.dept || ''), { align: 'left' });
  set(1, 2, '报销日期：' + (state.date || ''), { align: 'right' });
  /* 行3-4 表头 */
  set(2, 0, '摘要', { align: 'center' });
  set(2, 3, '金额', { align: 'center' });
  set(2, 11, '科目', { align: 'center' });
  set(2, 12, '单据\n张数', { align: 'center', align2: { wrapText: true } });
  DIGIT_LABELS.forEach(function (l, i) { set(3, 3 + i, l, { align: 'center' }); });
  /* 明细 */
  rows.forEach(function (row, li) {
    var r = 4 + li;
    set(r, 0, row.summary || '', { align: 'left', wrap: true });
    var vd = visibleDigits(row.digits);
    for (var d = 0; d < 8; d++) {
      var dv = vd[d] || '';
      set(r, 3 + d, dv === '' ? '' : (/^[0-9]+$/.test(dv) ? Number(dv) : dv), { align: 'center' });
    }
    set(r, 11, row.subject || '', { align: 'center', wrap: true });
    set(r, 12, row.count ? (/^[0-9]+$/.test(String(row.count)) ? Number(row.count) : row.count) : '', { align: 'center' });
  });
  /* 合计（仅本张） */
  var ftot = 0; rows.forEach(function (r) { ftot += num(r.amount); });
  set(9, 0, '合计（大写）：' + toCapital(ftot), { align: 'left' });
  var tds = visibleDigits(amountToDigits(ftot));
  for (var i = 0; i < 8; i++) {
    var tv = tds[i] || '';
    set(9, 3 + i, tv === '' ? '' : (/^[0-9]+$/.test(tv) ? Number(tv) : tv), { align: 'center' });
  }
  var fcnt = 0; rows.forEach(function (r) { fcnt += num(r.count); });
  set(9, 11, '单据 ' + Math.round(fcnt), { align: 'center' });
  /* 签字栏 */
  set(10, 0, '财会主管', { align: 'center' });
  set(10, 1, '记账', { align: 'center' });
  set(10, 2, '经办人', { align: 'center' });
  set(10, 3, '部门主管', { align: 'center' });
  set(10, 9, '复核', { align: 'center' });
  set(10, 12, '报销人', { align: 'center' });
  /* 合并 */
  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: 12 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } },
    { s: { r: 1, c: 2 }, e: { r: 1, c: 12 } },
    { s: { r: 2, c: 0 }, e: { r: 3, c: 2 } },
    { s: { r: 2, c: 3 }, e: { r: 2, c: 10 } },
    { s: { r: 2, c: 11 }, e: { r: 3, c: 11 } },
    { s: { r: 2, c: 12 }, e: { r: 3, c: 12 } },
    { s: { r: 9, c: 0 }, e: { r: 9, c: 2 } },
    { s: { r: 9, c: 11 }, e: { r: 9, c: 12 } },
    { s: { r: 10, c: 3 }, e: { r: 10, c: 8 } },
    { s: { r: 10, c: 9 }, e: { r: 10, c: 11 } }
  ];
  /* 每行明细「摘要」跨 3 列合并 */
  rows.forEach(function (_, li) { ws['!merges'].push({ s: { r: 4 + li, c: 0 }, e: { r: 4 + li, c: 2 } }); });
  /* 列宽（按原表样比例，单位：字符宽） */
  ws['!cols'] = COL_PCT.map(function (w) { return { wch: w }; });
  /* 行高（pt） */
  ws['!rows'] = [
    { hpt: 30 }, { hpt: 22 }, { hpt: 20 }, { hpt: 18 },
    { hpt: 26 }, { hpt: 26 }, { hpt: 26 }, { hpt: 26 }, { hpt: 26 },
    { hpt: 26 }, { hpt: 34 }
  ];
  ws['!pageSetup'] = { orientation: 'portrait', fitToWidth: 1, fitToHeight: 0, paperSize: 9 };
  ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 10, c: 12 } });
  return ws;
}

/* 导出前校验：返回仍有问题的明细行（已上传发票但识别为待补填 / 无金额 / 无摘要）。
   纯手工填写（无 invId）的行不强制，允许用户完全手填。 */
function problemRows() {
  var bad = [];
  state.rows.forEach(function (r, i) {
    if (!r.invId) return;                 // 手工行不强制
    var rec = state.invoices.filter(function (x) { return x.id === r.invId; })[0];
    var badRow = false;
    if (rec && rec.status === 'todo') badRow = true;        // 识别待补填
    if (!(num(r.amount) > 0)) badRow = true;                // 无金额
    if (!r.summary) badRow = true;                          // 无摘要
    if (badRow) bad.push({ i: i, rec: rec });
  });
  return bad;
}
/* 导出前检查：有待补填行则标红、滚动到首个问题处并阻止导出；返回 true 表示可继续 */
function validateBeforeExport() {
  var bad = problemRows();
  document.querySelectorAll('.bx-row-bad').forEach(function (el) { el.classList.remove('bx-row-bad'); });
  if (!bad.length) return true;
  bad.forEach(function (b) {
    var tr = document.querySelector('tr.bx-row[data-r="' + b.i + '"]');
    if (tr) tr.classList.add('bx-row-bad');
  });
  var first = bad[0];
  var ftr = document.querySelector('tr.bx-row[data-r="' + first.i + '"]');
  if (ftr) ftr.scrollIntoView({ behavior: 'smooth', block: 'center' });
  toast('有 ' + bad.length + ' 行待补填（已标红），请补全后再导出');
  return false;
}

/* Excel 导出（真 .xlsx，含样式）；XLSX_STYLE 不可用时回退 HTML .xls */
function exportExcel() {
  if (!validateBeforeExport()) return;
  var forms = formsFromRows();
  if (!forms.length) { toast('没有可导出的报销单'); return; }
  loadLib('xlsxStyle').then(function (ok) {
    var X = window.XLSX_STYLE;
    if (!ok || !X) { excelFallback(forms); return; }
    var wb = X.utils.book_new();
    forms.forEach(function (rows, k) {
      var ws = buildSheetForForm(rows, '报销单' + (k + 1));
      if (ws) X.utils.book_append_sheet(wb, ws, '报销单' + (k + 1));
    });
    var out = X.write(wb, { bookType: 'xlsx', type: 'array', cellStyles: true, bookSST: true });
    var blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = '费用报销单_' + (state.date || cnDate()) + '.xlsx';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { try { URL.revokeObjectURL(url); } catch (e) {} }, 1500);
    markExported();
    toast('已导出 Excel（' + (forms.length > 1 ? forms.length + ' 张报销单，每个工作表一张' : 'A4 竖版版式') + '）');
  }).catch(function () { excelFallback(forms); });
}
/* 兼容回退：旧 HTML .xls（多张堆叠） */
function excelFallback(forms) {
  var mso = '<!--[if gte mso 9]><xml><ExcelWorkbook><ExcelWorksheets><ExcelWorksheet>' +
    '<Name>费用报销单</Name><PageSetup><x:PaperSizeIndex>9</x:PaperSizeIndex>' +
    '<x:Orientation>Portrait</x:Orientation><x:FitWidth>1</x:FitWidth><x:FitHeight>0</x:FitHeight>' +
    '</PageSetup><Selected/></ExcelWorksheet></ExcelWorksheets></ExcelWorkbook></xml><![endif]-->';
  var parts = forms.map(function (rows) { return oneFormTableHtml(rows, 720); });
  var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" ' +
    'xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="utf-8"><title>费用报销单</title>' + mso + '</head><body>' +
    parts.join('<div style="height:18px"></div>') + '</body></html>';
  downloadDataUrl('data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(html),
    '费用报销单_' + (state.date || cnDate()) + '.xls');
  markExported();
  toast('已导出 Excel（兼容模式，多张堆叠）');
}
function exportWord() {
  if (!validateBeforeExport()) return;
  var forms = formsFromRows();
  if (!forms.length) { toast('没有可导出的报销单'); return; }
  var parts = forms.map(function (rows) { return officeHtml(rows); });
  downloadDataUrl('data:application/msword;charset=utf-8,' + encodeURIComponent(parts.join('<br>')),
    '费用报销单_' + (state.date || cnDate()) + '.doc');
  markExported();
  toast('已导出 Word（' + (forms.length > 1 ? forms.length + ' 张报销单' : '版式与预览一致') + '）');
}

/* 出图：临时进入「捕获模式」（藏按钮、去输入框底色），拍完恢复 */
function captureSheet() {
  return loadLib('html2canvas').then(function (ok) {
    if (!ok) throw new Error('出图组件不可用（网络受限）');
    var target = $('#sheet');
    document.body.classList.add('capturing');
    return new Promise(function (res) {
      setTimeout(function () {
        window.html2canvas(target, { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false })
          .then(function (c) { res(c); }, function (e) { res(null); });
      }, 60);
    }).then(function (c) {
      document.body.classList.remove('capturing');
      if (!c) throw new Error('出图失败');
      return c;
    }, function (e) {
      document.body.classList.remove('capturing');
      throw e;
    });
  });
}
function exportImage(kind) {
  if (!validateBeforeExport()) return;
  captureSheet().then(function (canvas) {
    var name = '费用报销单_' + (state.date || cnDate());
    if (kind === 'png') downloadDataUrl(canvas.toDataURL('image/png'), name + '.png');
    else downloadDataUrl(canvas.toDataURL('image/jpeg', 0.95), name + '.jpg');
    markExported();
    toast('已导出 ' + (kind === 'png' ? 'PNG' : 'JPG'));
  }).catch(function (e) {
    toast((e && e.message ? e.message : '出图失败') + '，改用打印窗口（可选“另存为 PDF”）');
    window.print();
  });
}
/* PDF：每 2 张报销单排在一个 A4 竖版页面的上下半页，等比（contain）缩放并略缩，保证整张表完整显示 */
function exportPdf() {
  if (!validateBeforeExport()) return;
  var forms = formsFromRows();
  if (!forms.length) { toast('没有可导出的报销单'); return; }
  var fixedPx = 794;
  var scaleDown = 0.97;
  Promise.all(forms.map(function (rows) { return renderFormToCanvas(rows, fixedPx); })).then(function (cans) {
    return loadLib('jspdf').then(function (ok) {
      if (!ok) throw new Error('PDF 组件不可用');
      var X = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
      var pdf = new X('p', 'pt', 'a4');             // A4 竖版 595.28 × 841.89
      var pw = pdf.internal.pageSize.getWidth();
      var ph = pdf.internal.pageSize.getHeight();
      var m = 12;                                    // 四周留白（pt）
      var perPage = 2;
      var slotH = (ph - 2 * m) / perPage;            // 半页高
      var slotW = pw - 2 * m;
      for (var i = 0; i < cans.length; i += perPage) {
        if (i > 0) pdf.addPage();
        for (var j = 0; j < perPage && i + j < cans.length; j++) {
          var c = cans[i + j];
          var s = Math.min(slotW / c.width, slotH / c.height) * scaleDown;  // 等比缩放，整张表完整
          var cw = c.width * s, ch = c.height * s;
          var cx = m + (slotW - cw) / 2;
          var cy = m + j * slotH + (slotH - ch) / 2;
          pdf.addImage(c.toDataURL('image/png'), 'PNG', cx, cy, cw, ch);
        }
      }
      markExported();
      pdf.save('费用报销单_' + (state.date || cnDate()) + '.pdf');
      toast('已导出 PDF（A4 竖版，每页上下两张、完整不裁切）');
    });
  }).catch(function (e) {
    toast((e && e.message ? e.message : '导出失败') + '，改用打印窗口，请选“另存为 PDF”');
    doPrint();
  });
}

/* 打印：专用打印区，每页上下两张报销单（与 PDF 一致的 2-up 版式） */
function doPrint() {
  if (!validateBeforeExport()) return;
  markExported();
  var pa = $('#printArea');
  if (pa) {
    var forms = formsFromRows();
    var html = '';
    var perPage = 2;
    for (var i = 0; i < forms.length; i += perPage) {
      html += '<div class="bx-print-page">';
      for (var j = 0; j < perPage && i + j < forms.length; j++) {
        html += '<table class="bx-table">' + buildFormHtml(forms[i + j], { screen: false, base: 0, k: 0 }) + '</table>';
      }
      html += '</div>';
    }
    pa.innerHTML = html;
  }
  window.print();
}

function downloadDataUrl(url, name) {
  var a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

/* ==================================================================
 * 十一、自检
 * ================================================================== */
function showDiagnose() {
  var rowsHtml = '';
  rowsHtml += diagRow('打开方式', location.protocol, IS_HTTP ? '可用本地随包引擎' : '浏览器禁止读本地引擎，只能走网络');
  Object.keys(LIBS).forEach(function (k) {
    var st = engineState[k];
    rowsHtml += diagRow(LIBS[k].label, st.status === 'ok' ? ('就绪 · ' + st.mirror) : (st.status === 'fail' ? '不可用 · ' + st.note : '加载中'), st.status === 'ok');
  });
  rowsHtml += diagRow('OCR 资源位置', ocrBase ? (ocrBase.corePath || '-') : '未就绪', !!ocrBase);
  rowsHtml += diagRow('已上传发票', state.invoices.length + ' 张', true);
  rowsHtml += diagRow('明细行数', state.rows.length + ' 行', true);
  rowsHtml += diagRow('合计金额', capital2(computeTotal()) + '（' + toCapital(computeTotal()) + '）', true);
  var box = $('#diagBody');
  box.innerHTML = rowsHtml +
    '<p style="margin-top:12px;color:#7a8699;font-size:12px">若某个组件显示「不可用」：请改用「双击启动」的 bat 打开（会读取随包 lib/，完全离线）；或联网后点左侧「重新加载」。</p>' +
    '<button id="clearReimb" class="btn danger sm" type="button" style="margin-top:14px">清除本地报销记录（解除拦截）</button>' +
    '<p style="margin-top:6px;color:#7a8699;font-size:12px">已上传过的发票会被记到本机浏览器，重复上传会拦截；点上面可清空记录。</p>';
  $('#clearReimb').addEventListener('click', function () {
    if (confirm('清除后，之前上传过的发票将不再被拦截（重复上传不会被拦）。确定清除？')) {
      try { localStorage.removeItem(REIMB_KEY); } catch (e) {}
      toast('已清除本地报销记录');
    }
  });
  $('#diagModal').classList.remove('hidden');
}
function diagRow(k, v, ok) {
  return '<div class="diag-row"><span>' + esc(k) + '</span><b class="' + (ok ? 'diag-ok' : 'diag-bad') + '">' + esc(v) + '</b></div>';
}

/* ==================================================================
 * 十二、启动
 * ================================================================== */
function bindGlobal() {
  $$('[data-export]').forEach(function (b) {
    b.addEventListener('click', function () {
      var k = b.dataset.export;
      if (k === 'excel') exportExcel();
      else if (k === 'word') exportWord();
      else if (k === 'pdf') exportPdf();
      else if (k === 'png') exportImage('png');
      else if (k === 'jpg') exportImage('jpg');
    });
  });
  $('[data-print]').addEventListener('click', function () { doPrint(); });
  $('[data-diagnose]').addEventListener('click', showDiagnose);
  $('[data-help]').addEventListener('click', function () { $('#helpModal').classList.remove('hidden'); });
  $$('[data-close-help]').forEach(function (b) {
    b.addEventListener('click', function () {
      $('#helpModal').classList.add('hidden');
      $('#diagModal').classList.add('hidden');
    });
  });
  [$('#helpModal'), $('#diagModal')].forEach(function (mo) {
    mo.addEventListener('click', function (e) { if (e.target === mo) mo.classList.add('hidden'); });
  });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { $('#helpModal').classList.add('hidden'); $('#diagModal').classList.add('hidden'); }
  });
}

function init() {
  state.date = cnDate();
  bindSheet();
  bindSidebar();
  bindGlobal();
  renderAll();
  updateEngineStatus();
  /* 后台预热引擎：本地可用时几乎瞬时，走网络时先把 6 个主脚本拉下来 */
  Object.keys(LIBS).forEach(function (k) { loadLib(k); });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

})();
