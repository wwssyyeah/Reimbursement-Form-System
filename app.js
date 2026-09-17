/* 报销单系统  v2.2 — 2026-09-16
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
  m = t.match(/发票号码[:：]?(\d{8,20})/); if (m) r.number = m[1];
  if (!r.number) { m = t.match(/票据号码[:：]?(\d{8,20})/); if (m) r.number = m[1]; }
  if (!r.number) { m = t.match(/号码[:：]?(\d{8,20})/); if (m) r.number = m[1]; }

  m = t.match(/开票日期[:：]?(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (m) r.date = m[1] + '-' + pad(m[2]) + '-' + pad(m[3]);
  if (!r.date) { m = t.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/); if (m) r.date = m[1] + '-' + pad(m[2]) + '-' + pad(m[3]); }
  if (!r.date) { m = t.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/); if (m) r.date = m[1] + '-' + pad(m[2]) + '-' + pad(m[3]); }

  /* 价税合计（小写）：跨过「（大写）壹仟…（小写）」取第一个两位小数 */
  m = t.match(/价税合计[^\d]{0,60}?(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/小写[^\d]{0,12}?(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/合计金额[^\d]{0,20}?(\d[\d,]*\.\d{2})/);
  if (!m) m = t.match(/\((?:小写)\)[^\d]{0,12}?(\d[\d,]*\.\d{2})/);
  if (m) r.total = m[1].replace(/,/g, '');
  if (!r.total) {
    /* 兜底：取全文中最大的两位小数（发票上通常是价税合计） */
    var cand = t.match(/\d[\d,]*\.\d{2}/g) || [];
    var max = 0;
    cand.forEach(function (x) { var v = num(x); if (v > max) max = v; });
    if (max > 0 && /发票|税额|合计/.test(t)) { r.total = String(max); r.totalGuessed = true; }
  }

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

  /* 候选公司名：整段文字含主体后缀即视为公司名（取完整文字项，避免停在公司名中间的「酒店」等字） */
  var compRe = /[一-龥A-Za-z0-9（）()·.\-]*?(?:公司|有限公司|有限责任公司|事务所|集团|酒店|旅行社|商行|中心|学校|医院|银行|合作社|厂|店|超市|科技|实业|工作室)/;
  var cands = [];
  its.forEach(function (it) {
    if (!compRe.test(it.s)) return;
    var name = it.s.replace(/^(名称|销方|卖方|销售方|购买方)?[:：]?\s*/, '').trim();
    if (name.length >= 4) cands.push({ name: name, x: it.x, y: it.y });
  });

  if (sx !== null) {
    /* 销售方在右侧栏：取 x 大于标签且最靠右的公司 */
    var right = cands.filter(function (c) { return c.x > sx - 6; });
    if (right.length) res.seller = right[right.length - 1].name;
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
  /* 摘要列不自动填，保留可编辑文本框供手工录入 */
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

async function processPdf(file, rec) {
  var ok = await loadLib('pdfjs');
  if (!ok) throw new Error('PDF 解析组件不可用（网络受限），请手动填写');
  await loadLib('jsqr');                       // 确保二维码解码库就绪
  var pl = window.pdfjsLib;
  if (pdfWorkerUrl) pl.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;
  var buf = await readAsArrayBuffer(file);
  var pdf = await pl.getDocument({ data: buf, isEvalSupported: false }).promise;
  var texts = [], qr = null, pdfItems = null;
  for (var p = 1; p <= pdf.numPages; p++) {
    var page = await pdf.getPage(p);
    var content = await page.getTextContent();
    var str = content.items.map(function (it) { return it.str; }).join(' ');
    if (flat(str).length < 30) {
      var viewport = page.getViewport({ scale: 2 });
      var canvas = document.createElement('canvas');
      canvas.width = viewport.width; canvas.height = viewport.height;
      var ctx = canvas.getContext('2d');
      ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: ctx, viewport: viewport }).promise;
      if (!qr) qr = decodeQrCanvas(canvas);     // 先扫二维码
      try { str = await ocrImage(canvas.toDataURL('image/png'), rec); } catch (e) { str = ''; }
    } else {
      if (!pdfItems) pdfItems = content.items;   // 保留文字层坐标，供版面感知解析（定位销售方/货物名）
      if (!qr) {                                // 文字够多也单独渲染扫二维码（二维码独立于文字层）
        try {
          var vp = page.getViewport({ scale: 1.5 });
          var cv = document.createElement('canvas'); cv.width = vp.width; cv.height = vp.height;
          var cctx = cv.getContext('2d'); cctx.fillStyle = '#fff'; cctx.fillRect(0, 0, cv.width, cv.height);
          await page.render({ canvasContext: cctx, viewport: vp }).promise;
          qr = decodeQrCanvas(cv);
        } catch (e) {}
      }
    }
    texts.push(str);
    if (qr) break;
  }
  return { text: texts.join('\n'), items: pdfItems, qr: qr };
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

async function handleFiles(fileList) {
  var files = Array.prototype.slice.call(fileList);
  if (!files.length) { toast('没有选到文件'); return; }
  var accept = ['pdf', 'jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif', 'doc', 'docx', 'xls', 'xlsx', 'csv'];
  var jobs = [], skipped = [];
  files.forEach(function (f) {
    var ext = extOf(f.name);
    if (accept.indexOf(ext) < 0) { skipped.push(f.name); return; }
    var rec = { id: uid(), fileName: f.name, ext: ext, thumb: '', status: 'ing', error: '排队中', rawText: '', progress: 0, qrOk: false, file: f };
    state.invoices.push(rec);
    var row = takeBlankRow();
    row.invId = rec.id;
    row.summary = '';   // 先占行，摘要留空供手工填写
    row.count = 1;
    jobs.push({ file: f, rec: rec, row: row, ext: ext });
  });
  renderAll();
  if (skipped.length) toast('忽略不支持的文件：' + skipped.join('、'));
  if (!jobs.length) return;
  toast('已接收 ' + jobs.length + ' 张，开始识别…（首张需先准备 OCR 组件）');
  for (var i = 0; i < jobs.length; i++) await processOne(jobs[i]);
  renderAll();
  var okN = state.invoices.filter(function (r) { return r.status === 'ok'; }).length;
  toast('识别完成：成功 ' + okN + ' 张，其余 ' + (state.invoices.length - okN) + ' 张可在表格里手工补填');
}

async function processOne(job) {
  var rec = job.rec, ext = job.ext;
  rec.status = 'ing'; rec.error = '识别中…';
  renderListThrottled();
  try {
    var text = '', qr = null, pdfRes = null;
    if (['jpg', 'jpeg', 'png', 'bmp', 'webp', 'gif'].indexOf(ext) >= 0) {
      rec.thumb = await readAsDataURL(job.file);
      renderList();
      var qrP = decodeQr(rec.thumb);            // 二维码优先（机器读取，最准）
      text = await ocrImage(rec.thumb, rec);    // OCR 补「销售方 / 项目名称」
      qr = await qrP;
    } else if (ext === 'pdf') {
      pdfRes = await processPdf(job.file, rec);
      text = pdfRes.text || ''; qr = pdfRes.qr;
    } else if (['xls', 'xlsx', 'csv'].indexOf(ext) >= 0) {
      text = await processExcel(job.file);
    } else if (['doc', 'docx'].indexOf(ext) >= 0) {
      text = await processWord(job.file);
    }
    rec.rawText = text || '';
    if (qr) rec.qrOk = true;
    var parsed = parseInvoice(text || '');
    /* 电子发票文字层：用坐标定位销售方/货物名（远准于 OCR，且纯本地） */
    if (pdfRes && pdfRes.items && pdfRes.items.length) {
      var lay = parsePdfLayout(pdfRes.items);
      if (lay.seller) parsed.seller = lay.seller;
      if (lay.item) parsed.item = lay.item;
    }
    var data = mergeInvoice(parsed, qr);
    data.vat = detectVat(text, qr);           // 科目：专票/普票（机器优先）
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
function renderAll() { renderList(); renderSheet(); updateTotals(); }

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

function renderSheet() {
  var t = $('#bxTable');
  if (!t) return;
  var rows = state.rows;
  var html = '';

  html += '<colgroup>';
  COL_PCT.forEach(function (w) { html += '<col style="width:' + w + '%">'; });
  html += '</colgroup>';

  /* 行1 标题 */
  html += '<tr><td class="bx-title" colspan="13">费  用  报  销  单</td></tr>';

  /* 行2 部门 | 报销日期 */
  html += '<tr>' +
    '<td class="bx-h" colspan="2">部门：<input id="f-dept" class="bx-in bx-line" type="text" value="' + esc(state.dept) + '"></td>' +
    '<td class="bx-h" colspan="11">报销日期：<input id="f-date" class="bx-in" type="text" value="' + esc(state.date) + '" placeholder="按发票开票日期"></td>' +
    '</tr>';

  /* 行3-4 表头 */
  html += '<tr>' +
    '<td class="bx-th" colspan="3" rowspan="2">摘要</td>' +
    '<td class="bx-th" colspan="8">金额</td>' +
    '<td class="bx-th" rowspan="2">科目</td>' +
    '<td class="bx-th" rowspan="2">单据<br>张数</td>' +
    '</tr><tr>';
  DIGIT_LABELS.forEach(function (l) { html += '<td class="bx-ths">' + l + '</td>'; });
  html += '</tr>';

  /* 行5.. 明细（全部为实体行，可直接输入） */
  rows.forEach(function (row, i) {
    html += '<tr class="bx-row">' +
      '<td class="bx-summary" colspan="3">' +
      '<button class="bx-del screen-only" type="button" data-r="' + i + '" title="删除本行">✕</button>' +
      '<input class="bx-in" type="text" data-r="' + i + '" data-k="summary" value="' + esc(row.summary) + '" placeholder="摘要（可手填）">' +
      '</td>';
    var vd = visibleDigits(row.digits);
    for (var d = 0; d < 8; d++) {
      html += '<td class="bx-digit"><input class="bx-in" type="text" maxlength="1" inputmode="numeric" data-r="' + i + '" data-d="' + d + '" value="' + esc(vd[d] || '') + '"></td>';
    }
    html += '<td>' + '<input class="bx-in" type="text" data-r="' + i + '" data-k="subject" value="' + esc(row.subject) + '">' + '</td>';
    html += '<td><input class="bx-in bx-incount" type="text" maxlength="3" inputmode="numeric" data-r="' + i + '" data-k="count" value="' + esc(row.count) + '"></td>';
    html += '</tr>';
  });

  /* 合计（大写）行 */
  var tot = computeTotal();
  var td = amountToDigits(tot);
  var vtd = visibleDigits(td);
  html += '<tr class="bx-total">' +
    '<td class="bx-tl" colspan="3">合计（大写）：<b id="t-capital">' + toCapital(tot) + '</b></td>';
  vtd.forEach(function (v, i) {
    html += '<td class="bx-digit bx-tot"><span id="tot-d-' + i + '">' + (v || '') + '</span></td>';
  });
  html += '<td class="bx-tc" colspan="2">单据 <b id="t-totalcount">' + totalCount() + '</b></td></tr>';

  /* 签字栏（严格按原文件列位：E:J 为「部门主管」合并区）；每个栏位可手填姓名 */
  function sgColspan(label, key, cs) {
    return '<td colspan="' + cs + '"><div class="sg-cell"><span class="sg-l">' + label + '</span>' +
      '<input class="bx-in sg-in" type="text" data-s="' + key + '" value="' + esc(state.sign[key]) + '"></div></td>';
  }
  html += '<tr class="bx-sign">' +
    sgColspan('财会主管', 'caik', 1) +
    sgColspan('记账', 'jizhang', 1) +
    sgColspan('经办人', 'jingban', 1) +
    sgColspan('部门主管', 'bumen', 6) +
    '<td colspan="3"><div class="sg-cell sg-inline"><span class="sg-l">复核</span>' +
      '<input class="bx-in sg-in" type="text" data-s="fuhe" value="' + esc(state.sign.fuhe) + '"></div></td>' +
    sgColspan('报销人', 'baoxiao', 1) +
    '</tr>';

  t.innerHTML = html;

  var fd = $('#f-dept'); if (fd) fd.addEventListener('input', function () { state.dept = fd.value; });
  var ft = $('#f-date'); if (ft) ft.addEventListener('input', function () { state.date = ft.value; });
  var et = $('#emptyTip'); if (et) et.classList.toggle('hidden', hasData());
}

function updateTotals() {
  var tot = computeTotal();
  var cap = $('#t-capital'); if (cap) cap.textContent = toCapital(tot);
  var ds = visibleDigits(amountToDigits(tot));
  for (var i = 0; i < 8; i++) {
    var el = document.getElementById('tot-d-' + i);
    if (el) el.textContent = ds[i] || '';
  }
  var tc = $('#t-totalcount'); if (tc) tc.textContent = totalCount();
  var et = $('#emptyTip'); if (et) et.classList.toggle('hidden', hasData());
}

/* ==================================================================
 * 九、编辑交互
 * ================================================================== */
function bindSheet() {
  var t = $('#bxTable');
  if (!t) return;

  t.addEventListener('input', function (e) {
    var el = e.target;
    if (!el || !el.dataset) return;
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
      if (r && confirm('删除这张发票？对应的明细行也会清空。')) {
        state.invoices = state.invoices.filter(function (x) { return x.id !== id; });
        state.rows.forEach(function (row) { if (row.invId === id) { row.invId = null; row.summary = ''; row.subject = ''; row.count = ''; row.amount = 0; row.digits = emptyDigits(); } });
        renderAll();
      }
    }
  });

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
  if (!row) { row = takeBlankRow(); row.invId = id; row.count = 1; }
  await processOne({ file: rec.file, rec: rec, row: row, ext: rec.ext });
  toast('「' + rec.fileName + '」已重新识别');
}

/* ==================================================================
 * 十、导出
 * ================================================================== */
/* 生成「干净」表格：去掉操作按钮，输入框换成纯文字（避免导出图出现多余横线） */
function cleanTableClone() {
  var src = $('#bxTable');
  if (!src) return null;
  var t = src.cloneNode(true);
  Array.prototype.forEach.call(t.querySelectorAll('.screen-only'), function (n) {
    if (n.parentNode) n.parentNode.removeChild(n);
  });
  Array.prototype.forEach.call(t.querySelectorAll('input,select,textarea'), function (el) {
    var sp = document.createElement('span');
    sp.className = (el.className || '') + ' bx-val';
    sp.textContent = el.value || '';
    if (el.id === 'f-dept') sp.className += ' bx-line-val';
    if (el.parentNode) el.parentNode.replaceChild(sp, el);
  });
  return t;
}
/* 给每个格子写内联样式，保证 Word/Excel 里也有框线和居中 */
function inlineize(t) {
  Array.prototype.forEach.call(t.querySelectorAll('td'), function (td) {
    var cls = ' ' + (td.className || '') + ' ';
    var st = 'border:1px solid #000000;padding:4px 3px;vertical-align:middle;';
    if (/ bx-title /.test(cls)) st += 'font-size:18pt;font-weight:bold;text-align:center;padding:8px 0;letter-spacing:2pt;';
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
function officeHtml() {
  var t = cleanTableClone();
  if (!t) return '';
  inlineize(t);
  return '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns="http://www.w3.org/TR/REC-html40">' +
    '<head><meta charset="utf-8"><title>费用报销单</title></head><body>' +
    '<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;font-family:Microsoft YaHei,sans-serif;font-size:11pt">' +
    t.innerHTML + '</table></body></html>';
}
function exportExcel() {
  downloadDataUrl('data:application/vnd.ms-excel;charset=utf-8,' + encodeURIComponent(officeHtml()),
    '费用报销单_' + (state.date || cnDate()) + '.xls');
  toast('已导出 Excel（版式与预览一致）');
}
function exportWord() {
  downloadDataUrl('data:application/msword;charset=utf-8,' + encodeURIComponent(officeHtml()),
    '费用报销单_' + (state.date || cnDate()) + '.doc');
  toast('已导出 Word（版式与预览一致）');
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
  captureSheet().then(function (canvas) {
    var name = '费用报销单_' + (state.date || cnDate());
    if (kind === 'png') downloadDataUrl(canvas.toDataURL('image/png'), name + '.png');
    else downloadDataUrl(canvas.toDataURL('image/jpeg', 0.95), name + '.jpg');
    toast('已导出 ' + (kind === 'png' ? 'PNG' : 'JPG'));
  }).catch(function (e) {
    toast((e && e.message ? e.message : '出图失败') + '，改用打印窗口（可选“另存为 PDF”）');
    window.print();
  });
}
function exportPdf() {
  captureSheet().then(function (canvas) {
    return loadLib('jspdf').then(function (ok) {
      if (!ok) throw new Error('PDF 组件不可用');
      var js = (window.jspdf && window.jspdf.jsPDF) ? window.jspdf.jsPDF : window.jsPDF;
      var pdf = new js({ orientation: 'p', unit: 'pt', format: [canvas.width, canvas.height] });
      pdf.addImage(canvas.toDataURL('image/png'), 'PNG', 0, 0, canvas.width, canvas.height);
      pdf.save('费用报销单_' + (state.date || cnDate()) + '.pdf');
      toast('已导出 PDF');
    });
  }).catch(function (e) {
    toast((e && e.message ? e.message : '导出失败') + '，改用打印窗口，请选“另存为 PDF”');
    window.print();
  });
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
    '<p style="margin-top:12px;color:#7a8699;font-size:12px">若某个组件显示「不可用」：请改用「双击启动」的 bat 打开（会读取随包 lib/，完全离线）；或联网后点左侧「重新加载」。</p>';
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
  $('[data-print]').addEventListener('click', function () { window.print(); });
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
