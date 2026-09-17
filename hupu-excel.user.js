// ==UserScript==
// @name         虎扑 Excel · 摸鱼模式
// @namespace    https://bbs.hupu.com/
// @version      1.1.0
// @author       you
// @license      MIT
// @description  把 bbs.hupu.com 伪装成 Excel 工作簿：读页面自带的 $$data / __NEXT_DATA__ 渲染成带行号列标的表格，支持点选单元格、公式栏、多工作表、翻页；右上角 ⚙ 打开设置面板，Esc 一键切回原页面。
// @match        *://*.hupu.com/*
// @match        *://hupu.com/*
// @grant        GM_getValue
// @grant        GM_setValue
// @run-at       document-start
// @noframes
// ==/UserScript==

/*
 * ── 这个脚本干了什么 ────────────────────────────────────────────────────────
 *
 * reference/nga-excel.js（NGA 优化摸鱼体验）的 Excel 模式是「保留原生 DOM + 纯 CSS
 * 改造成表格」；reference/nga-codex.user.js 则是「读结构化数据 + 自己渲染」。
 * 虎扑的页面结构和 NGA 完全不同，这里取两者之长：
 *
 *   1. 数据读取
 *      - 帖子页（/123456789.html、/123456789-2.html）是 Next.js SSR，
 *        页面里带 <script id="__NEXT_DATA__">，标题 / 正文 / 作者 / 楼层 /
 *        点亮数 / 分页全在里面，比 DOM 权威得多 —— 首选它。
 *      - 版块列表页（/bxj、/bxj-2、/topic-daily …）与虎扑社区首页是另一套
 *        React（bbs-pc-svc）渲染的，没有 __NEXT_DATA__，只能解析 DOM。
 *      - 两条通道最后都被归一化成同一份 {sheets:[{cols,rows}]} 模型。
 *
 *   2. 伪装
 *      - 原生 DOM 只是 display:none 藏起来（不移除），站点自己的 JS 照常跑，
 *        链接、登录态、埋点都不受影响。
 *      - 站内链接走「软导航」：fetch 回目标页的 HTML、解析出数据后在同一个
 *        文档里重画，浏览器不换文档 —— 所以切版面 / 进帖子 / 翻页都不会
 *        闪一下原生页面（详见第 10 节）。
 *      - 自己渲染一整窗 Excel：标题栏 / 选项卡 / 功能区 / 编辑栏 / 行列标 /
 *        冻结窗格 / 工作表标签 / 状态栏，行号列标用 Excel 的真实规则（A..Z、AA..）。
 *      - 支持点选单元格（名称框 + 编辑栏联动）、方向键移动、回车打开链接、
 *        Ctrl+PageUp/PageDown 切换工作表、Alt+←/→ 翻页。
 *
 *   3. 逃生
 *      - Alt+E 开关整个 Excel 模式（写进 GM 存储）。
 *      - Esc 立即切回原页面（不写存储），再按一次回来 —— 老板键。
 *
 *   4. 配置
 *      - 全部配置项都在右上角 ⚙ 打开的「Excel 选项」面板里（不注册油猴菜单项）。
 * ──────────────────────────────────────────────────────────────────────────
 */

(function () {
  'use strict';

  /*
   * 启动时序。
   *
   * 关于「切页面时会闪一下原生页面」：hupu 的热缓存页面从 document 创建到第一次
   * 绘制只有 ~30ms，而浏览器扩展（油猴）的 document-start 注入要先走一趟后台
   * service worker，通常晚几十到几百 ms —— 这段时间里页面已经画出来了，脚本
   * 还没开始跑，页面里做什么都拦不住。这里把这几个时刻记下来，控制台会打印，
   * 设置面板「关于」里也能看到，用来判断到底晚在哪一段。
   */
  const BOOT = {
    script: (typeof performance !== 'undefined' ? Math.round(performance.now()) : 0),
    hide: -1, shell: -1, domReady: -1, firstPaint: -1, fcp: -1
  };

  /*
   * 「这次加载直接显示原生页面」的一次性信号。
   *
   * 软导航（见第 10 节）只换 Excel 视图、不换物理 DOM，所以软导航之后按 Esc
   * 切回原页面会看到上一次整页加载的旧内容。为了老板键靠得住，这种时候我们会
   * 重新加载当前地址，并在这里留个标记：新文档启动时直接进原生模式，
   * 而不是又变回 Excel。标记只消费一次。
   */
  let START_NATIVE = false;
  try {
    START_NATIVE = sessionStorage.getItem('hx.native') === '1';
    if (START_NATIVE) sessionStorage.removeItem('hx.native');
  } catch (e) { /* 无 sessionStorage 时忽略 */ }

  /* ============================== 1. 配置与存储 ============================== */

  const DEFAULTS = {
    enabled: true,        // 是否开启 Excel 模式
    theme: 'office',      // office | tencent | wps
    book: '工作簿1',       // 工作簿名（会写进浏览器标签标题；留空则用页面标题）
    showAccount: true,    // 标题栏右侧的账号区
    showUrl: false,       // 表格里的「路径」列（默认关闭，更清爽）
    freezeHeader: true,   // 冻结列标行 / 行号列
    showImages: true,     // 帖子正文里的图片
    imgMaxW: 260,         // 缩略图宽度上限（px）
    imgMaxH: 170,         // 缩略图高度上限（px）
    imgHoverZoom: true    // 鼠标悬停浮出大图
  };

  function storeGet(key, fallback) {
    try {
      const raw = typeof GM_getValue === 'function' ? GM_getValue(key, null) : null;
      return raw == null ? fallback : JSON.parse(raw);
    } catch (e) {
      return fallback;
    }
  }

  function storeSet(key, value) {
    try {
      if (typeof GM_setValue === 'function') GM_setValue(key, JSON.stringify(value));
    } catch (e) { /* 忽略 */ }
  }

  const CFG = Object.assign({}, DEFAULTS, storeGet('hx.cfg', {}));
  function saveCfg() { storeSet('hx.cfg', CFG); }

  /* ============================== 2. 小工具 ============================== */

  /*
   * DOC 是「当前要解析的文档」。
   *
   * 正常情况下它就是 document（原生页面）；做软导航（见第 10 节）时会被临时
   * 换成 fetch 回来、用 DOMParser 解析出的那份文档，这样数据读取那套函数不用
   * 改签名，就能直接读「新页面」的 $$data / __NEXT_DATA__ / DOM。
   */
  let DOC = document;
  const $$ = (sel, root) => (root || DOC).querySelector(sel);
  const $$$ = (sel, root) => Array.prototype.slice.call((root || DOC).querySelectorAll(sel));
  /** 取第一个匹配元素（找不到返回 null，比 querySelector()[0] 安全） */
  const one = (sel, root) => (root || DOC).querySelector(sel);

  function el(tag, cls, text) {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text != null) node.textContent = text;
    return node;
  }

  /** 取文本（折叠空白），NGA 那套 txt() 的同款语义 */
  function txt(node) {
    return node ? String(node.textContent || '').replace(/\s+/g, ' ').trim() : '';
  }

  function attr(node, name) {
    return node && node.getAttribute ? (node.getAttribute(name) || '') : '';
  }

  /** 页面数据拼 HTML 前一律转义 */
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[c]));
  }

  function num(v) {
    const n = Number(String(v == null ? '' : v).replace(/[^\d.-]/g, ''));
    return isNaN(n) ? 0 : n;
  }

  /** 千分位；非数字原样返回 */
  function fmtNum(v) {
    const raw = String(v == null ? '' : v).trim();
    if (!raw) return '';
    if (!/^-?[\d,]+(\.\d+)?$/.test(raw)) return raw;
    const n = Number(raw.replace(/,/g, ''));
    return isNaN(n) ? raw : n.toLocaleString('en-US');
  }

  function pad2(n) { return n < 10 ? '0' + n : String(n); }

  /** 13 位时间戳 / 字符串 → 2026-09-17 10:34:00 */
  function fmtTime(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number' || /^\d{10,13}$/.test(String(v))) {
      let ms = Number(v);
      if (String(Math.trunc(ms)).length === 10) ms *= 1000;
      const d = new Date(ms);
      if (isNaN(d.getTime())) return String(v);
      return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) +
        ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
    }
    return String(v).replace(/T/, ' ').replace(/\.\d+Z?$/, '');
  }

  /** 相对/绝对地址 → 绝对地址 */
  function abs(u) {
    if (!u) return '';
    u = String(u).trim();
    if (/^data:/i.test(u)) return '';
    if (/^https?:\/\//i.test(u)) return u;
    if (u.indexOf('//') === 0) return location.protocol + u;
    if (u.charAt(0) === '/') return location.origin + u;
    if (u.charAt(0) === '#') return '';
    try { return new URL(u, location.href).href; } catch (e) { return u; }
  }

  /** 从帖子链接里抽 tid：优先 /123456789.html，兼容绝对地址与 ?tid= */
  function tidOf(href) {
    const s = String(href || '');
    let m = s.match(/\/(\d+)(?:-\d+)?\.html/);
    if (m) return m[1];
    m = s.match(/[?&]tid=(\d+)/);
    if (m) return m[1];
    m = s.match(/(\d{5,})/);
    return m ? m[1] : '';
  }

  /** 0 → A、25 → Z、26 → AA */
  function colName(i) {
    let s = '';
    i = i + 1;
    while (i > 0) {
      const m = (i - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      i = Math.floor((i - 1) / 26);
    }
    return s;
  }

  function clamp(v, min, max) { return v < min ? min : (v > max ? max : v); }

  /* ============================== 3. 数据读取 ============================== */

  /** 读取 Next.js 内联数据（帖子页 SSR 时存在） */
  function readNextData() {
    const node = DOC.getElementById('__NEXT_DATA__');
    if (!node) return null;
    try { return JSON.parse(node.textContent); } catch (e) { return null; }
  }

  /** 路由判定：首页 / 版块列表 / 帖子 / 其它 */
  function route() {
    const path = location.pathname.replace(/\/+$/, '') || '/';
    if (path === '/') return { kind: 'home' };
    let m = path.match(/^\/(\d+)(?:-(\d+))?\.html$/);
    if (m) return { kind: 'thread', tid: m[1], page: m[2] ? +m[2] : 1 };
    m = path.match(/^\/([\w-]+?)(?:-(\d+))?$/);
    if (m) return { kind: 'list', slug: m[1], page: m[2] ? +m[2] : 1 };
    return { kind: 'other' };
  }

  /** 面包屑：社区 » 步行街 » 步行街主干道 */
  function crumbs() {
    return crumbList().map(c => c.title);
  }

  /** 面包屑（带链接）：编辑栏里那一串可以点的「社区 » 步行街 » 动物萌宠区」 */
  function crumbList() {
    const out = [];
    $$$('.bbs-sl-web-bread-crumbs a, .hp-pc-breadcrumb a').forEach(a => {
      const t = txt(a);
      if (!t || (out.length && out[out.length - 1].title === t)) return;
      out.push({ title: t, url: abs(attr(a, 'href')) });
    });
    return out;
  }

  function pageTitleText() {
    return (DOC.title || '').replace(/\s*[-—|]\s*虎扑.*$/, '').trim();
  }

  /**
   * 帖子里的图片：缩略图用 data-zoom 记下大图地址，鼠标悬停时浮出。
   * 虎扑的图床 URL 带缩放参数（x-oss-process=image/resize,w_800 / imageMogr2/…），
   * 去掉 query 就是原图。
   */
  function bigImageUrl(thumb, origin) {
    const o = abs(origin || '');
    if (o) return o;
    const u = abs(thumb || '');
    if (!u) return '';
    if (/x-oss-process=|imageMogr2|imageView2/i.test(u)) return u.split('?')[0];
    return u;
  }

  /**
   * 从一段 HTML 里抽出「可放进单元格」的内容。
   *
   * 虎扑正文的 HTML 很简单（p / img / br / a / blockquote），直接白名单式重建：
   *   - img  → 单元格里的浮动图片（Excel 里粘图的既视感）
   *   - br / 块级标签结尾 → 换行
   *   - a    → 蓝色下划线超链接
   *   - 其余标签 → 丢掉标签保留文字
   */
  function contentBox(html) {
    /*
     * 返回值必须是 Element，不能是 DocumentFragment：
     * DocumentFragment 是一次性的 —— appendChild 之后它的子节点就被搬走了，
     * 再 append 一次就是空的。而我们的表格会反复重渲染（切工作表、改设置、
     * 翻页…），用 Fragment 会表现为「切一次工作表正文就全空了」。
     */
    const box = document.createElement('div');
    box.className = 'hx-rich';
    if (!html) return box;
    const BLOCK = /^(P|DIV|LI|UL|OL|H[1-6]|BLOCKQUOTE|SECTION|ARTICLE|TABLE|TR|FIGURE|PRE|HR)$/;

    let doc;
    try {
      doc = new DOMParser().parseFromString('<div id="hx-src">' + html + '</div>', 'text/html');
    } catch (e) { return box; }
    const src = doc.getElementById('hx-src');
    if (!src) return box;

    const walk = (node, out, depth) => {
      if (depth > 24) return;
      const kids = node.childNodes;
      for (let i = 0; i < kids.length; i++) {
        const n = kids[i];
        if (n.nodeType === 3) {
          const t = String(n.nodeValue || '').replace(/[\t\r\n]+/g, ' ').replace(/ {2,}/g, ' ');
          if (t.trim()) out.appendChild(document.createTextNode(t));
          continue;
        }
        if (n.nodeType !== 1) continue;
        const tag = n.tagName;

        if (tag === 'IMG') {
          const thumb = abs(attr(n, 'src') || attr(n, 'data-src') || attr(n, 'data-origin'));
          const zoom = bigImageUrl(thumb, attr(n, 'data-origin') || attr(n, 'data-zoom'));
          if (thumb || zoom) {
            if (CFG.showImages) {
              const im = document.createElement('img');
              im.className = 'hx-img';
              im.src = thumb || zoom;
              im.dataset.zoom = zoom;
              im.loading = 'lazy';
              im.referrerPolicy = 'no-referrer';
              out.appendChild(im);
            } else {
              const a = document.createElement('a');
              a.className = 'hx-link';
              a.href = zoom || thumb;
              a.target = '_blank';
              a.rel = 'noreferrer';
              a.dataset.zoom = zoom || thumb;
              a.textContent = '【图片】';
              out.appendChild(a);
            }
            out.appendChild(document.createElement('br'));
          }
          continue;
        }
        if (tag === 'BR') { out.appendChild(document.createElement('br')); continue; }
        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'NOSCRIPT') continue;
        if (tag === 'IFRAME' || tag === 'VIDEO' || tag === 'EMBED' || tag === 'SOURCE') {
          const url = abs(attr(n, 'src'));
          if (url) {
            const a = document.createElement('a');
            a.className = 'hx-link';
            a.href = url;
            a.target = '_blank';
            a.rel = 'noreferrer';
            a.textContent = '【视频/外链】' + url;
            out.appendChild(a);
            out.appendChild(document.createElement('br'));
          }
          continue;
        }
        if (tag === 'A') {
          const href = abs(attr(n, 'href'));
          const box = document.createElement('a');
          box.className = 'hx-link';
          if (href) { box.href = href; box.target = '_blank'; box.rel = 'noreferrer'; }
          walk(n, box, depth + 1);
          if (txt(box) || box.querySelector('img')) out.appendChild(box);
          continue;
        }
        if (tag === 'BLOCKQUOTE') {
          const q = document.createElement('div');
          q.className = 'hx-quote';
          walk(n, q, depth + 1);
          out.appendChild(q);
          continue;
        }
        walk(n, out, depth + 1);
        if (BLOCK.test(tag)) out.appendChild(document.createElement('br'));
      }
    };

    walk(src, box, 0);
    // 去掉结尾多余的换行
    while (box.lastChild && box.lastChild.nodeName === 'BR') box.removeChild(box.lastChild);
    return box;
  }

  /** HTML → 纯文本（给编辑栏 / title 提示用） */
  function htmlToText(html) {
    return txt(contentBox(html));
  }

  /* ---------- 3.0 内联数据 window.$$data ---------- */

  /*
   * 虎扑的列表页和首页是「服务端渲染 + 内联 JSON」：
   *     <script>window.$$data={"topic":{...}}</script>
   * 这份 JSON 比 DOM 完整得多：
   *   - 首页 DOM 只渲染了 60 条帖子，JSON 里有 70 条；
   *   - 列表页 DOM 没有「点亮数」，JSON 里有 lights；
   *   - 首页 JSON 里有 13 个分类各自的全部板块（255 个）和「我的关注」follow。
   * 所以优先读它，DOM 只作兜底。帖子页是 Next.js，走 __NEXT_DATA__。
   */
  let PAGE_DATA;
  function readPageData() {
    if (PAGE_DATA !== undefined) return PAGE_DATA;
    PAGE_DATA = null;
    const scripts = DOC.querySelectorAll('script:not([src])');
    for (let i = 0; i < scripts.length; i++) {
      const text = scripts[i].textContent || '';
      const at = text.indexOf('window.$$data');
      if (at < 0) continue;
      const start = text.indexOf('{', at);
      const end = text.lastIndexOf('}');
      if (start < 0 || end <= start) continue;
      try { PAGE_DATA = JSON.parse(text.slice(start, end + 1)); } catch (e) { PAGE_DATA = null; }
      break;
    }
    return PAGE_DATA;
  }

  /** 登录状态（首页 / 列表页的内联数据里有 isLogin） */
  function isLoggedIn() {
    const data = readPageData();
    if (!data) return false;
    if (data.isLogin) return true;
    return !!(data.pageData && data.pageData.isLogin);
  }

  /** 当前用户的登录信息：优先读页面顶部的账号 DOM（站点 JS 异步填进去的） */
  function readAccount() {
    const box = one('.hp-topLogin-info');
    if (!box) return null;
    const link = one('a[href*="my.hupu.com"]', box) || one('a[href*="i.hupu.com"]', box);
    const nameEl = one('[class*="name"]', box) || link;
    let name = txt(nameEl) || txt(box);
    name = name.replace(/^(你好|Hi|hi)[，,、\s]*/, '').replace(/\s+/g, ' ').trim();
    if (!name || /^登录|注册|登录后的世界/.test(name)) return null;
    const img = one('img', box);
    return {
      name: name.slice(0, 20),
      avatar: img ? abs(attr(img, 'src') || attr(img, 'data-src')) : '',
      url: link ? abs(attr(link, 'href')) : ''
    };
  }

  /** 板块路径显示成 /topic-daily 这种短路径，而不是完整 URL */
  function shortPath(url) {
    const s = String(url || '');
    if (!s) return '';
    const m = s.match(/^https?:\/\/[^/]+(\/.*)$/);
    return m ? m[1] : s;
  }

  /* ---------- 3.1 版块列表页（/bxj、/bxj-2、/topic-daily …） ---------- */

  /** 板块表：我的关注 / 热门专区 / … */
  function boardSheet(name, list, srcName) {
    const cols = [
      { label: '序号', width: 56, align: 'right' },
      { label: '板块', width: 178 },
      { label: '分类', width: 108 },
      { label: '热度', width: 84, align: 'right' }
    ];
    if (srcName) cols.push({ label: '来源', width: 84 });
    cols.push({ label: '路径', width: 170, url: true });

    return {
      name: name,
      cols: cols,
      rows: (list || []).map((f, i) => {
        const cells = [
          { text: String(i + 1) },
          { text: f.name || '', href: abs(f.url), plain: f.name || '' },
          { text: f.cat || '', href: f.catHref ? abs(f.catHref) : '' },
          { text: f.countText || (f.count ? fmtNum(f.count) : ''), raw: f.count || 0 }
        ];
        if (srcName) cells.push({ text: f.src || '' });
        cells.push({ text: shortPath(f.url) });
        return { cells: cells };
      })
    };
  }

  /** 用内联数据渲染列表页 */
  function listFromData(r, data, tp, th) {
    const forum = (tp.topic && tp.topic.name) || pageTitleText();
    const crumbList = (tp.breadCrumb || []).filter(b => b.title).map(b => ({ title: b.title, url: abs(b.url) }));
    const sheets = [{
      name: forum,
      cols: [
        { label: '编号', width: 92 },
        { label: '标题', width: 600 },
        { label: '回复', width: 64, align: 'right' },
        { label: '浏览', width: 76, align: 'right' },
        { label: '点亮', width: 64, align: 'right' },
        { label: '作者', width: 132 },
        { label: '最后回复', width: 150 },
        { label: '路径', width: 150, url: true }
      ],
      rows: (th.list || []).map(t => ({
        cells: [
          { text: t.tid || '' },
          { text: t.title || '', href: abs(t.url), plain: t.title || '' },
          { text: fmtNum(t.replies), raw: num(t.replies) },
          { text: fmtNum(t.read), raw: num(t.read) },
          { text: fmtNum(t.lights), raw: num(t.lights) },
          { text: (t.author && t.author.puname) || '', href: t.author && t.author.url },
          { text: fmtTime(t.repliedAt || t.createdAt) },
          { text: shortPath(t.url) }
        ]
      }))
    }];

    // 左侧栏的「我的关注 / 话题广场」（跟首页用的是同一套数据）
    navSheets().forEach(sheet => sheets.push(sheet));

    const admins = (data.admins && data.admins.adminList) || [];
    if (admins.length) {
      sheets.push({
        name: '版主',
        cols: [
          { label: '序号', width: 56, align: 'right' },
          { label: '角色', width: 110 },
          { label: '用户名', width: 160 },
          { label: '主页', width: 320, url: true }
        ],
        rows: admins.map((a, i) => ({
          cells: [
            { text: String(i + 1) },
            { text: a.roleName || '' },
            { text: a.userName || '', href: a.url },
            { text: shortPath(a.url) }
          ]
        }))
      });
    }

    const slug = r.slug || '';
    return {
      title: forum,
      crumbs: crumbList.map(c => c.title),
      crumbList: crumbList,
      sheetName: forum,
      pager: {
        current: num(th.current) || r.page || 1,
        total: num(th.total) || 1,
        href: n => n <= 1 ? '/' + slug : '/' + slug + '-' + n
      },
      sheets: sheets
    };
  }

  /** 内联数据缺失时，退回 DOM 解析 */
  function listFromDom(r) {
    const rows = [];
    $$$('li.bbs-sl-web-post-body').forEach(li => {
      const a = one('.post-title a.p-title', li) || one('.post-title a', li);
      if (!a) return;
      const href = abs(attr(a, 'href'));
      const datum = txt(one('.post-datum', li)).split('/');
      const authorA = one('.post-auth a', li);
      const pageNums = $$$('.page-icon a', li).map(x => txt(x)).filter(Boolean);
      rows.push({
        cells: [
          { text: tidOf(attr(a, 'href')) },
          {
            text: txt(a),
            href: href,
            plain: txt(a) + (pageNums.length ? '（共 ' + pageNums[pageNums.length - 1] + ' 页）' : '')
          },
          { text: fmtNum(datum[0]), raw: num(datum[0]) },
          { text: fmtNum(datum[1]), raw: num(datum[1]) },
          { text: '' },
          { text: txt(authorA), href: attr(authorA, 'href') ? abs(attr(authorA, 'href')) : '' },
          { text: txt(one('.post-time', li)) },
          { text: shortPath(attr(a, 'href')) }
        ]
      });
    });

    if (!rows.length) return null;

    const items = $$$('.hupu-rc-pagination-item');
    let total = 1;
    items.forEach(li => {
      const m = String(li.className).match(/hupu-rc-pagination-item-(\d+)/);
      if (m) total = Math.max(total, +m[1]);
    });
    const current = num(txt(one('.hupu-rc-pagination-item-active'))) || r.page || 1;

    const forum = txt(one('.bbs-sl-web-intro-detail-title')).replace(/^#/, '') ||
      (crumbs().slice(-1)[0] || '') || pageTitleText() || '帖子列表';

    const sheets = [{
      name: forum,
      cols: [
        { label: '编号', width: 92 },
        { label: '标题', width: 600 },
        { label: '回复', width: 64, align: 'right' },
        { label: '浏览', width: 76, align: 'right' },
        { label: '点亮', width: 64, align: 'right' },
        { label: '作者', width: 132 },
        { label: '最后回复', width: 110 },
        { label: '路径', width: 150, url: true }
      ],
      rows: rows
    }];
    navSheets().forEach(sheet => sheets.push(sheet));

    return {
      title: forum,
      crumbs: crumbs(),
      crumbList: crumbList(),
      sheetName: forum,
      pager: {
        current: current,
        total: total,
        href: n => n <= 1 ? '/' + r.slug : '/' + r.slug + '-' + n
      },
      sheets: sheets
    };
  }

  function modelList(r) {
    const data = readPageData();
    const tp = data && data.topic;
    const th = tp && tp.threads;
    if (th && th.list && th.list.length) {
      try { return listFromData(r, data, tp, th); } catch (e) { /* 退回 DOM */ }
    }
    // 分类页（/all-gambia、/all-sports…）：$$data 在 pageData 里，不在 topic 里
    if (data && !tp) {
      const zone = modelZone(data);
      if (zone) return zone;
    }
    return listFromDom(r);
  }

  /**
   * 分类页 —— 面包屑里点「步行街」「综合体育」这种。
   *
   * 它既不是版面页也不是首页：$$data 在 pageData 里，
   *   pageData.category = { name, url, hot / topics: [子板块…] }
   *   pageData.threads  = 该分类下的帖子流（70 条）
   * 以前只认 data.topic.threads，于是整页落到「没有数据」的兜底表。
   * 这里把它摊成「分类帖子流 + 分类下的板块」两张表，跟版面页观感一致。
   */
  function modelZone(data) {
    const pd = data && data.pageData;
    const cat = pd && pd.category;
    if (!cat) return null;

    const threads = (pd.threads || []).filter(t => t && t.url);
    // topics 不是每次都下发，hot 一定有；两个都当子板块列表用
    const boardList = (cat.topics && cat.topics.length) ? cat.topics : (cat.hot || []);
    const boards = boardList.filter(b => b && b.name && b.url);
    if (!threads.length && !boards.length) return null;

    const name = cat.name || pageTitleText() || '分类';
    const sheets = [];
    if (threads.length) sheets.push(feedSheet(name, threads));
    if (boards.length) {
      sheets.push(boardSheet(name + ' · 板块', boards.map(b => ({
        name: b.name,
        url: b.url,
        cat: name,
        count: b.count,
        countText: b.countText,
        cateId: b.cateId || cat.cateId
      })), false));
    }
    navSheets().forEach(s => sheets.push(s));

    const crumbList = [
      { title: '社区', url: 'https://bbs.hupu.com/' },
      { title: name, url: abs(cat.url) }
    ];
    return {
      title: name,
      crumbs: crumbList.map(c => c.title),
      crumbList: crumbList,
      sheetName: sheets[0].name,
      pager: null,
      sheets: sheets
    };
  }

  /* ---------- 3.2 帖子页（__NEXT_DATA__ 优先，DOM 兜底） ---------- */

  function threadFromNext(next, r) {
    const det = next && next.props && next.props.pageProps && next.props.pageProps.detail;
    if (!det || !det.thread) return null;
    const th = det.thread;
    const rep = det.replies || {};

    let content = th.content || '';
    if (!htmlToText(content) && th.format) {
      try {
        const f = JSON.parse(th.format);
        content = (f && (f.htmlV3 || (f.jsonV3 && JSON.stringify(f.jsonV3)))) || content;
      } catch (e) { /* 忽略 */ }
    }

    const floors = [];
    // 楼主
    floors.push({
      cells: [
        { text: '楼主' },
        { text: (th.author && th.author.puname) || '', href: th.author && th.author.url },
        { node: contentBox(content), plain: htmlToText(content) },
        { text: fmtNum(th.lights), raw: num(th.lights) },
        { text: fmtTime(th.createdAt) },
        { text: [th.location, th.client].filter(Boolean).join(' · ') }
      ]
    });

    const offset = Math.max(0, (num(rep.current) - 1) * num(rep.size || 20));
    (rep.list || []).forEach((p, i) => {
      floors.push({
        cells: [
          { text: offset + i + 1 },
          { text: (p.author && p.author.puname) || '', href: p.author && p.author.url },
          { node: contentBox(p.content), plain: htmlToText(p.content) },
          { text: fmtNum(p.count), raw: num(p.count) },
          { text: fmtTime(p.createdAt) },
          { text: [p.location, p.client].filter(Boolean).join(' · ') }
        ]
      });
    });

    const sheets = [{
      name: '全部楼层',
      cols: [
        { label: '楼层', width: 62 },
        { label: '作者', width: 132 },
        { label: '内容', width: 720, wrap: true, content: true },
        { label: '点亮', width: 62, align: 'right' },
        { label: '时间', width: 150 },
        { label: '属地/端', width: 118 }
      ],
      rows: floors
    }];

    // 侧栏的热帖 / 最新帖，顺手做成另外两张工作表
    const relSheet = (name, list) => ({
      name: name,
      cols: [
        { label: '编号', width: 92 },
        { label: '标题', width: 600 },
        { label: '回复', width: 64, align: 'right' },
        { label: '浏览', width: 76, align: 'right' },
        { label: '作者', width: 132 },
        { label: '时间', width: 150 }
      ],
      rows: (list || []).map(t => ({
        cells: [
          { text: t.tid },
          { text: t.title, href: abs(t.url || ('/' + t.tid + '.html')), plain: t.title },
          { text: fmtNum(t.replies), raw: num(t.replies) },
          { text: fmtNum(t.read), raw: num(t.read) },
          { text: (t.author && t.author.puname) || '', href: t.author && t.author.url },
          { text: fmtTime(t.createdAt) }
        ]
      }))
    });
    if ((det.hot || []).length) sheets.push(relSheet('热帖', det.hot));
    if ((det.latest || []).length) sheets.push(relSheet('最新', det.latest));

    return {
      title: th.title || pageTitleText(),
      crumbs: (det.breadCrumb || []).map(b => b.title).filter(Boolean),
      crumbList: (det.breadCrumb || []).filter(b => b.title).map(b => ({ title: b.title, url: abs(b.url) })),
      sheetName: '全部楼层',
      pager: {
        current: num(rep.current) || r.page || 1,
        total: num(rep.total) || 1,
        href: n => n <= 1 ? '/' + r.tid + '.html' : '/' + r.tid + '-' + n + '.html'
      },
      sheets: sheets
    };
  }

  /** __NEXT_DATA__ 缺失时的 DOM 兜底（结构变了也不至于整页空白） */
  function threadFromDom(r) {
    const titleNode = one('[class*="post-info-bottom-title"]') || one('h1');
    const title = txt(titleNode) || pageTitleText();
    const contentNode = one('.thread-content-detail');
    if (!title && !contentNode) return null;

    const nameNode = one('[class*="post-user-comp-info-top-name"]');
    const timeNode = one('[class*="post-user-comp-info-top-time"]');
    const locNode = one('[class*="user-location"]');

    const floors = [{
      cells: [
        { text: '楼主' },
        { text: txt(nameNode), href: abs(attr(nameNode, 'href')) },
        { node: contentBox(contentNode ? contentNode.innerHTML : ''), plain: txt(contentNode) },
        { text: '' },
        { text: txt(timeNode).replace('发布于', '') },
        { text: txt(locNode).replace('发布于', '') }
      ]
    }];

    $$$('.post-reply-list-wrapper').forEach((wrap, i) => {
      const n = one('[class*="user-info-top-name"]', wrap);
      const t = one('[class*="user-info-top-time"]', wrap);
      const lo = one('[class*="user-info-user-location"]', wrap);
      const c = one('.thread-content-detail', wrap);
      const light = one('.post-reply-list-operate .todo-list', wrap);
      floors.push({
        cells: [
          { text: i + 1 },
          { text: txt(n), href: abs(attr(n, 'href')) },
          { node: contentBox(c ? c.innerHTML : ''), plain: txt(c) },
          { text: txt(light).replace(/[^\d]/g, '') },
          { text: txt(t) },
          { text: txt(lo).replace('发布于', '') }
        ]
      });
    });

    const pages = $$$('.hupu-rc-pagination-item').map(li => {
      const m = String(li.className).match(/hupu-rc-pagination-item-(\d+)/);
      return m ? +m[1] : 0;
    });
    const total = pages.length ? Math.max.apply(null, pages) : 1;

    return {
      title: title,
      crumbs: crumbs(),
      crumbList: crumbList(),
      sheetName: '全部楼层',
      pager: { current: r.page || 1, total: total, href: n => n <= 1 ? '/' + r.tid + '.html' : '/' + r.tid + '-' + n + '.html' },
      sheets: [{
        name: '全部楼层',
        cols: [
          { label: '楼层', width: 62 },
          { label: '作者', width: 132 },
          { label: '内容', width: 720, wrap: true, content: true },
          { label: '点亮', width: 62, align: 'right' },
          { label: '时间', width: 150 },
          { label: '属地/端', width: 118 }
        ],
        rows: floors
      }]
    };
  }

  /* ---------- 3.3 左侧栏：我的关注 / 话题广场（首页与版块页共用） ---------- */

  /** "1.4w" / "7240" / "89.2w" → 数字（用于排序） */
  function heatValue(s) {
    const m = String(s == null ? '' : s).match(/([\d.]+)\s*(w|W|万)?/);
    if (!m) return 0;
    let v = parseFloat(m[1]);
    if (isNaN(v)) return 0;
    if (m[2]) v *= 10000;
    return Math.round(v);
  }

  /**
   * 页面左侧栏「话题广场」区的完整板块名单。
   *
   * 注意：内联 JSON 里 categories[].hot 只带了每个分类的前几个板块（共 61 个），
   * 完整名单（13 类 / 255 个）在 DOM 的弹层里。所以名单以 DOM 为准，
   * 热度数字再用 JSON 覆盖成精确值。
   */
  function boardsFromDom() {
    const cats = [];
    const seen = Object.create(null);
    $$$('.hu-pc-navigation-topic-type-item').forEach(box => {
      const titleEl = one('.hu-pc-navigation-topic-type-title', box);
      const cat = txt(titleEl) || '其它';
      const href = abs(attr(titleEl, 'href'));
      const items = [];
      $$$('.topic-item', box).forEach(it => {
        const nameEl = one('.topic-item-name', it);
        const name = txt(nameEl) || txt(it);
        const url = abs(attr(it, 'href'));
        if (!name || !url || seen[url]) return;
        seen[url] = 1;
        const heatRaw = txt(one('.topic-item-heat', it));
        items.push({ name: name, url: url, count: heatValue(heatRaw), countText: heatRaw, cat: cat });
      });
      if (items.length) cats.push({ name: cat, href: href, items: items });
    });
    return cats;
  }

  /**
   * 左侧栏的数据：DOM 出名单 + 内联 JSON 出精确数字。
   * 版块列表页（$$data.topic）和首页（$$data.pageData）的子结构是一样的。
   */
  function navData() {
    const data = readPageData() || {};
    const src = data.topic || data.pageData || {};
    const catJson = src.categories || [];

    const catById = Object.create(null);
    catJson.forEach(c => { catById[c.cateId] = c.name; });

    const heat = Object.create(null);
    (src.hot || []).forEach(f => { if (f && f.url) heat[f.url] = f; });
    catJson.forEach(c => (c.hot || []).forEach(f => { if (f && f.url) heat[f.url] = f; }));

    let cats = boardsFromDom();
    if (!cats.length) {
      cats = catJson.map(c => ({
        name: c.name,
        href: abs(c.url),
        items: (c.hot || []).map(f => ({
          name: f.name, url: abs(f.url), count: num(f.count), countText: f.countText || '', cat: c.name
        }))
      })).filter(c => c.items.length);
    }
    cats.forEach(c => c.items.forEach(f => {
      const j = heat[f.url];
      if (j) { f.count = num(j.count); f.countText = j.countText || f.countText; }
    }));

    const catOfUrl = Object.create(null);
    cats.forEach(c => c.items.forEach(f => { catOfUrl[f.url] = c.name; }));

    return { cats: cats, catById: catById, catOfUrl: catOfUrl, heat: heat, follow: src.follow || [], hot: src.hot || [] };
  }

  /** 板块表：序号 / 板块 / 分类 / 热度 (/ 来源) (/ 路径) */
  function boardSheet(name, list, srcName) {
    const cols = [
      { label: '序号', width: 56, align: 'right' },
      { label: '板块', width: 178 },
      { label: '分类', width: 108 },
      { label: '热度', width: 84, align: 'right' }
    ];
    if (srcName) cols.push({ label: '来源', width: 84 });
    cols.push({ label: '路径', width: 170, url: true });

    return {
      name: name,
      cols: cols,
      rows: (list || []).map((f, i) => {
        const cells = [
          { text: String(i + 1) },
          { text: f.name || '', href: abs(f.url), plain: f.name || '' },
          { text: f.cat || '', href: f.catHref ? abs(f.catHref) : '' },
          { text: f.countText || (f.count ? fmtNum(f.count) : ''), raw: num(f.count) }
        ];
        if (srcName) cells.push({ text: f.src || '' });
        cells.push({ text: shortPath(f.url) });
        return { cells: cells };
      })
    };
  }

  /** 「我的关注」工作表：关注的板块在前，热门在后；未登录时首行给个说明 */
  function buildBoard(f, nav, src) {
    const j = nav.heat[f.url];
    return {
      name: f.name,
      url: abs(f.url),
      count: j ? num(j.count) : num(f.count || 0),
      countText: (j && j.countText) || f.countText || '',
      cat: nav.catById[f.cateId] || nav.catOfUrl[f.url] || f.cat || '',
      src: src
    };
  }

  /** 我关注的板块：内联 JSON 的 follow 优先，其次 DOM，都没有就空数组 */
  function collectFollows(nav) {
    const out = [];
    const seen = Object.create(null);
    (nav.follow || []).forEach(f => {
      const url = abs(f.url);
      if (!url || !f.name || seen[url]) return;
      seen[url] = 1;
      out.push(buildBoard({ name: f.name, url: url, count: f.count, countText: f.countText, cateId: f.cateId }, nav, '我的关注'));
    });
    if (!out.length) {
      $$$('.hu-pc-navigation-my-focus-item a[href]').forEach(a => {
        const name = txt(a);
        const url = abs(attr(a, 'href'));
        if (!name || !url || /登录|注册/.test(name) || seen[url]) return;
        seen[url] = 1;
        out.push(buildBoard({ name: name, url: url }, nav, '我的关注'));
      });
    }
    return out;
  }

  /** 热门板块（内联 JSON 的 hot 优先，其次按话题广场热度取前 N） */
  function collectHot(nav, exclude) {
    const out = [];
    const seen = Object.create(null);
    (exclude || []).forEach(f => { seen[f.url] = 1; });
    const push = (f) => {
      const url = abs(f.url);
      if (!url || !f.name || seen[url]) return;
      seen[url] = 1;
      out.push(buildBoard({ name: f.name, url: url, count: f.count, countText: f.countText, cateId: f.cateId }, nav, '热门'));
    };
    (nav.hot || []).forEach(push);
    if (!out.length) {
      const all = [];
      nav.cats.forEach(c => c.items.forEach(f => all.push(f)));
      all.sort((a, b) => num(b.count) - num(a.count)).slice(0, 24).forEach(push);
    }
    return out;
  }

  /**
   * 「我的关注」工作表。
   *
   * 登录并且拿到关注列表时，这里**只放我关注的板块** —— 以前把热门也拼进来，
   * 结果用户看到的是一堆自己没关注的版面（「我的关注」名不副实）。
   * 关注列表拿不到（未登录 / 页面没给）时，才用热门兜底，并在首行说明。
   */
  function followSheet(nav) {
    const follows = collectFollows(nav);
    if (follows.length) return boardSheet('我的关注', follows, false);

    const hot = collectHot(nav, []);
    if (!hot.length) return null;
    const sheet = boardSheet('我的关注', hot, false);
    sheet.rows.unshift({
      group: isLoggedIn()
        ? '没读到你的关注列表 · 下面是热门板块'
        : '未登录 · 登录后这里显示你关注的板块；下面是热门板块'
    });
    return sheet;
  }

  /** 「话题广场」工作表：13 个分类成段，段标题行 + 该分类下的板块 */
  function zoneSheet(cats) {
    const rows = [];
    cats.forEach(c => {
      const list = c.items.slice().sort((a, b) => num(b.count) - num(a.count));
      const total = list.reduce((s, x) => s + num(x.count), 0);
      rows.push({
        group: c.name,
        groupNote: list.length + ' 个板块 · 热度 ' + fmtNum(total),
        href: c.href
      });
      list.forEach((f, i) => rows.push({
        cells: [
          { text: String(i + 1) },
          { text: f.name, href: f.url, plain: f.name },
          { text: f.countText || fmtNum(f.count), raw: num(f.count) },
          { text: shortPath(f.url) }
        ]
      }));
    });
    return {
      name: '话题广场',
      cols: [
        { label: '序号', width: 56, align: 'right' },
        { label: '板块', width: 200 },
        { label: '热度', width: 90, align: 'right' },
        { label: '路径', width: 180, url: true }
      ],
      rows: rows
    };
  }

  /** 左侧栏的两张表（首页、版块列表页共用） */
  function navSheets() {
    const nav = navData();
    const out = [];
    const follows = collectFollows(nav);
    const follow = followSheet(nav);
    if (follow) out.push(follow);
    // 有关注列表时，「我的关注」里只有关注，热门另开一张表
    if (follows.length) {
      const hot = collectHot(nav, follows);
      if (hot.length) out.push(boardSheet('热门专区', hot, false));
    }
    if (nav.cats.length) out.push(zoneSheet(nav.cats));
    return out;
  }

  /** 帖子流工作表 */
  function feedSheet(name, threads) {
    return {
      name: name,
      cols: [
        { label: '版块', width: 130 },
        { label: '编号', width: 92 },
        { label: '标题', width: 660 },
        { label: '点亮', width: 64, align: 'right' },
        { label: '回复', width: 64, align: 'right' },
        { label: '路径', width: 140, url: true }
      ],
      rows: threads.map(t => ({
        cells: [
          { text: (t.topic && t.topic.name) || t.section || '', href: t.topic && abs(t.topic.url) },
          { text: t.tid || '' },
          { text: t.title || '', href: abs(t.url), plain: t.title || '' },
          { text: fmtNum(t.lights), raw: num(t.lights) },
          { text: fmtNum(t.replies), raw: num(t.replies) },
          { text: shortPath(t.url) }
        ]
      }))
    };
  }

  /** 首页推荐流（内联数据没有时的 DOM 兜底） */
  function threadsFromDom() {
    const posts = [];
    let section = '';
    $$$('.list-item-wrap').forEach(wrap => {
      const titleNode = one('.list-title', wrap);
      if (titleNode) { section = txt(titleNode); return; }
      $$$('.list-item', wrap).forEach(item => {
        const titleEl = one('.t-title', item);
        if (!titleEl) return;
        const a = titleEl.closest('a') || one('.t-info a', item);
        const boardEl = one('.t-label a', item);
        posts.push({
          tid: tidOf(attr(a, 'href')),
          title: txt(titleEl),
          url: abs(attr(a, 'href')),
          lights: txt(one('.t-lights', item)).replace(/[^\d]/g, ''),
          replies: txt(one('.t-replies', item)).replace(/[^\d]/g, ''),
          section: section || txt(boardEl),
          topic: { name: txt(boardEl), url: abs(attr(boardEl, 'href')) }
        });
      });
    });
    return posts;
  }

  /**
   * 社区首页：底部表标签跟页面自己的分区对应 ——
   *   我的关注 | 话题广场 | 站内热帖
   */
  function modelHome() {
    const data = readPageData();
    const pd = data && data.pageData;
    const sheets = navSheets();

    let threads = (pd && pd.threads) || [];
    if (!threads.length) threads = threadsFromDom();
    if (threads.length) sheets.push(feedSheet('站内热帖', threads));

    if (!sheets.length) return null;
    return {
      title: '虎扑社区',
      crumbs: ['社区'],
      crumbList: [{ title: '社区', url: 'https://bbs.hupu.com/' }],
      sheetName: sheets[0].name,
      pager: null,
      sheets: sheets
    };
  }

  /** 统一入口：当前页面 → 模型 */
  function buildModel() {
    const r = route();
    if (r.kind === 'home') return modelHome();
    if (r.kind === 'thread') {
      return threadFromNext(readNextData(), r) || threadFromDom(r);
    }
    if (r.kind === 'list') return modelList(r);
    // 其它页面（个人中心、搜索…）尽量按列表页解析，但分页信息不可信，丢掉
    const m = modelList({ kind: 'list', slug: '', page: 1 });
    if (m) m.pager = null;
    return m;
  }

  /* ============================== 4. 外观（CSS） ============================== */

  /*
   * 全部样式都限定在 #hx-root 里，并用一段 scope reset 挡掉虎扑自己的全局样式
   * （虎扑是 CSS Modules + 少量全局 tag 选择器，不挡会串味）。
   * 主题通过 CSS 变量切换：office（默认，绿）/ tencent（蓝）/ wps（深蓝）。
   */
  const CSS = `
  html.hx-on, html.hx-on body { overflow: hidden !important; background: #fff !important; }
  /*
   * 隐藏原生页面：body 下除了我们自己的 #hx-root 全藏掉。
   *
   * 故意不写死 #container / #__next 这类容器 id —— 错误页、WAF 拦截页、
   * 以后改版的容器都可能是别的，漏掉一个就是一整段原生页面被看见
   * （「切页面时有时会看到原生页面」就是这么来的）。全藏并不影响功能：
   * 原生 DOM 还在内存里，站点自己的 JS 照跑，我们照样能读写它。
   */
  html.hx-on body > *:not(#hx-root),
  html.hx-on > *:not(head):not(body):not(#hx-root) { display: none !important; }
  /* 上面的规则还有一份不经过 scopeCss 的快速通道（HIDE_CSS），首帧用 */
  html.hx-peek #hx-root { display: none !important; }

  #hx-root {
    --accent: #217346;
    --accent-dark: #175c38;
    --accent-soft: #d3e5dc;
    --title-bg: #ffffff;
    --title-fg: #201f1e;
    --title-sub: #605e5c;
    --chrome-bg: #ffffff;
    --ribbon-bg: #ffffff;
    --border: #e1dfdd;
    --border-strong: #d4d4d4;
    --gridline: #ececec;
    --head-bg: #f9fafb;
    --head-fg: #5a5a5a;
    --link: #0563c1;
    --text: #1f1f1f;
    --muted: #777777;
    --img-max-w: 260px;
    --img-max-h: 170px;

    position: fixed; inset: 0; z-index: 2147483000;
    display: flex; flex-direction: column;
    background: #fff; color: var(--text);
    font-family: "Segoe UI", "Microsoft YaHei UI", "Microsoft YaHei", system-ui, sans-serif;
    font-size: 12px; line-height: 1.4;
    -webkit-font-smoothing: antialiased;
    text-align: left;
  }
  /* scope reset：只清会影响布局/外观的属性，颜色在各组件上单独写 */
  #hx-root, #hx-root * {
    box-sizing: border-box; margin: 0; padding: 0; border: 0;
    font: inherit; color: inherit; letter-spacing: normal; text-transform: none;
    vertical-align: baseline; float: none; list-style: none; text-decoration: none;
    min-width: 0; max-width: none; background-repeat: no-repeat;
  }
  #hx-root svg { display: block; width: 16px; height: 16px; fill: none; stroke: currentColor; stroke-width: 1.2; stroke-linecap: round; stroke-linejoin: round; }
  #hx-root a { color: inherit; }
  #hx-root img { max-width: none; }
  #hx-root input { color: inherit; }

  /* ---------- 主题：腾讯文档 ---------- */
  #hx-root.hx-t-tencent {
    --accent: #1e6fff; --accent-dark: #0b53d6; --accent-soft: #e2ecff;
    --border: #ebebeb; --border-strong: #e0e0e0; --gridline: #ebebeb;
    --head-bg: #f9fafb; --head-fg: #464d5a; --link: #1e6fff;
    font-family: -apple-system, "Helvetica Neue", Helvetica, "PingFang SC", "Microsoft YaHei", sans-serif;
  }
  #hx-root.hx-t-tencent .hx-ribbon { display: none; }
  #hx-root.hx-t-tencent .hx-titlebar { height: 40px; }
  #hx-root.hx-t-tencent .hx-tabs { height: 26px; border-bottom-color: #ebebeb; }
  #hx-root.hx-t-tencent .hx-tab { font-size: 12px; color: #646a73; }
  #hx-root.hx-t-tencent .hx-tab.file { background: none; color: #646a73; font-weight: 400; }
  #hx-root.hx-t-tencent .hx-tab.active { color: var(--accent); font-weight: 500; border-color: transparent; }

  /* ---------- 主题：WPS ---------- */
  #hx-root.hx-t-wps {
    --accent: #2b6cd4; --accent-dark: #1f56ad; --accent-soft: #dbe7fb;
    --title-bg: #2b6cd4; --title-fg: #ffffff; --title-sub: #cfdcf7;
    --border: #d3dded; --border-strong: #c2cfe4; --gridline: #e5e9f0;
    --head-bg: #f0f3f9; --head-fg: #5b6b85; --link: #1a53b8;
  }
  #hx-root.hx-t-wps .hx-titlebar { border-bottom: none; }
  #hx-root.hx-t-wps .hx-search { background: rgba(255,255,255,.14); border-color: rgba(255,255,255,.35); color: #fff; }
  #hx-root.hx-t-wps .hx-search::placeholder { color: rgba(255,255,255,.7); }
  #hx-root.hx-t-wps .hx-win, #hx-root.hx-t-wps .hx-acct { color: #fff; }
  #hx-root.hx-t-wps .hx-tab.file { background: #1f56ad; }

  /* ---------- 标题栏 ---------- */
  .hx-titlebar {
    flex: 0 0 auto; height: 44px; display: flex; align-items: center; gap: 6px;
    padding: 0 6px 0 10px; background: var(--title-bg); color: var(--title-fg);
    border-bottom: 1px solid var(--border);
  }
  .hx-logo { flex: 0 0 auto; width: 22px; height: 22px; }
  .hx-logo svg { width: 22px; height: 22px; }
  .hx-book { font-size: 13px; font-weight: 600; padding: 3px 8px; border-radius: 3px; white-space: nowrap; color: var(--title-fg); }
  .hx-book:hover { background: rgba(0,0,0,.05); }
  .hx-save { font-size: 11px; color: var(--title-sub); display: flex; align-items: center; gap: 5px; white-space: nowrap; }
  .hx-dot { width: 7px; height: 7px; border-radius: 50%; background: var(--accent); display: inline-block; }
  .hx-tbtn {
    width: 24px; height: 24px; display: flex; align-items: center; justify-content: center;
    border-radius: 3px; color: var(--title-sub); flex: 0 0 auto;
  }
  .hx-tbtn:hover { background: rgba(0,0,0,.06); }
  .hx-tbtn svg { width: 15px; height: 15px; }
  .hx-spacer { flex: 1 1 auto; }
  .hx-acct { display: flex; align-items: center; gap: 12px; color: var(--title-sub); font-size: 12px; white-space: nowrap; }
  .hx-acct .hx-chip { display: flex; align-items: center; gap: 4px; }
  .hx-acct .hx-chip svg { width: 14px; height: 14px; }
  .hx-avatar { width: 20px; height: 20px; border-radius: 50%; background: var(--accent); color: #fff; font-size: 11px; display: flex; align-items: center; justify-content: center; }
  .hx-search {
    flex: 0 0 auto; width: 200px; height: 26px; border: 1px solid var(--border-strong);
    border-radius: 2px; background: #fff; color: var(--text);
    padding: 0 8px; font-size: 12px; outline: none;
  }
  .hx-win { display: flex; align-items: center; color: var(--title-sub); margin-left: 6px; flex: 0 0 auto; }
  .hx-win i { width: 34px; height: 30px; display: flex; align-items: center; justify-content: center; font-style: normal; font-size: 12px; }
  .hx-win i:hover { background: rgba(0,0,0,.07); }
  .hx-win i.close:hover { background: #e81123; color: #fff; }

  /* ---------- 选项卡 ---------- */
  .hx-tabs {
    flex: 0 0 auto; height: 30px; display: flex; align-items: flex-end; gap: 1px;
    padding: 0 6px; background: var(--chrome-bg); border-bottom: 1px solid var(--border);
  }
  .hx-tab {
    height: 28px; line-height: 28px; padding: 0 12px; font-size: 12px; color: #444;
    border: 1px solid transparent; border-bottom: none; border-radius: 3px 3px 0 0;
    white-space: nowrap; flex: 0 0 auto;
  }
  .hx-tab:hover { background: rgba(0,0,0,.04); }
  .hx-tab.file { background: var(--accent); color: #fff; font-weight: 600; padding: 0 15px; }
  .hx-tab.active { background: #fff; color: var(--accent); font-weight: 600; border-color: var(--border); position: relative; }
  .hx-tab.active::after { content: ""; position: absolute; left: -1px; right: -1px; bottom: -1px; height: 2px; background: #fff; }
  .hx-tabs-right { margin-left: auto; display: flex; align-items: center; gap: 8px; color: #666; font-size: 12px; padding-bottom: 5px; }
  .hx-tabs-right svg { width: 14px; height: 14px; }

  /* ---------- 功能区 ---------- */
  .hx-ribbon {
    flex: 0 0 auto; height: 112px; display: flex; align-items: stretch;
    padding: 4px 6px 0; background: var(--ribbon-bg);
    border-bottom: 1px solid var(--border); overflow: hidden;
  }
  .hx-group { display: flex; flex-direction: column; justify-content: space-between; padding: 0 7px; border-right: 1px solid #eaeaea; flex: 0 0 auto; }
  .hx-group:last-child { border-right: none; }
  .hx-gbody { display: flex; align-items: flex-start; gap: 1px; flex: 1 1 auto; }
  .hx-gtitle { font-size: 10px; line-height: 14px; color: #9a9a9a; text-align: center; padding: 1px 0 2px; white-space: nowrap; }
  .hx-btn {
    display: flex; flex-direction: column; align-items: center; gap: 1px;
    min-width: 34px; max-width: 72px; padding: 3px 5px 1px; border: 1px solid transparent;
    border-radius: 3px; color: #3b3a39; cursor: default;
  }
  .hx-btn:hover { background: #f0f0f0; border-color: #e4e4e4; }
  .hx-btn .i { height: 22px; display: flex; align-items: center; justify-content: center; }
  .hx-btn .i svg { width: 17px; height: 17px; }
  .hx-btn .t { font-size: 10px; line-height: 14px; color: #555; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 100%; }
  .hx-btn.big { padding-top: 2px; }
  .hx-btn.big .i { height: 30px; }
  .hx-btn.big .i svg { width: 26px; height: 26px; }
  .hx-col2 { display: flex; flex-direction: column; gap: 1px; justify-content: center; }
  .hx-col2 .hx-btn { flex-direction: row; gap: 5px; min-width: 64px; padding: 2px 5px; }
  .hx-col2 .hx-btn .i { height: 16px; }
  .hx-col2 .hx-btn .i svg { width: 14px; height: 14px; }
  .hx-glyph { font-family: "Times New Roman", Times, serif; font-size: 15px; line-height: 1; }
  .hx-glyph.b { font-weight: 700; }
  .hx-glyph.i { font-style: italic; }
  .hx-glyph.u { text-decoration: underline; }
  .hx-glyph.s { text-decoration: line-through; }
  .hx-bar { display: block; width: 15px; height: 3px; border-radius: 1px; margin-top: 1px; }
  .hx-combo {
    display: flex; align-items: center; justify-content: space-between; gap: 6px;
    height: 22px; padding: 0 6px; border: 1px solid var(--border-strong); border-radius: 2px;
    background: #fff; font-size: 12px; color: #333; min-width: 96px;
  }
  .hx-combo svg { width: 10px; height: 10px; color: #888; }
  .hx-combo.wide { min-width: 134px; }
  .hx-numcell { display: flex; flex-direction: column; align-items: center; gap: 3px; }
  .hx-numrow { display: flex; align-items: center; gap: 4px; }
  .hx-numrow .hx-btn { min-width: 26px; padding: 2px 4px; }
  .hx-numrow .hx-btn .i svg { width: 15px; height: 15px; }
  .hx-numrow .hx-btn .t { font-size: 10px; }
  .hx-tiny { min-width: 18px; padding: 2px 2px; }
  .hx-tiny .i svg { width: 11px; height: 11px; }

  /* ---------- 编辑栏 ---------- */
  .hx-formulabar {
    flex: 0 0 auto; height: 26px; display: flex; align-items: stretch;
    background: #fff; border-bottom: 1px solid var(--border); font-size: 12px;
  }
  .hx-namebox { flex: 0 0 116px; width: 116px; line-height: 25px; padding: 0 8px; color: #333; border-right: 1px solid var(--border); overflow: hidden; white-space: nowrap; text-overflow: ellipsis; }
  .hx-fx { flex: 0 0 34px; text-align: center; line-height: 25px; font-family: Georgia, serif; font-style: italic; color: #8a8886; border-right: 1px solid var(--border); }
  .hx-formula { flex: 1 1 auto; line-height: 25px; padding: 0 8px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; color: #333; }

  /* ---------- 表格 ---------- */
  .hx-sheet { flex: 1 1 auto; overflow: auto; background: #fff; position: relative; }
  .hx-table { min-width: max-content; }
  .hx-tr { display: flex; align-items: stretch; min-height: 22px; }
  .hx-th { position: sticky; top: 0; z-index: 2; background: var(--head-bg); }
  .hx-corner {
    position: sticky; left: 0; z-index: 4; flex: 0 0 50px; width: 50px; height: 22px;
    background: var(--head-bg); border-right: 1px solid var(--border-strong); border-bottom: 1px solid var(--border-strong);
  }
  .hx-corner::after {
    content: ""; position: absolute; right: 3px; bottom: 3px;
    border-top: 6px solid transparent; border-left: 6px solid transparent;
    border-right: 6px solid #b8b8b8; border-bottom: 6px solid #b8b8b8;
    display: block;
  }
  .hx-coll {
    flex: 0 0 auto; height: 22px; line-height: 21px; text-align: center;
    background: var(--head-bg); color: var(--head-fg); font-size: 11px;
    border-right: 1px solid #e6e6e6; border-bottom: 1px solid var(--border-strong);
    font-family: Calibri, "Segoe UI", sans-serif; overflow: hidden;
  }
  .hx-rowhead {
    position: sticky; left: 0; z-index: 1; flex: 0 0 50px; width: 50px; min-height: 22px;
    background: var(--head-bg); color: var(--head-fg); text-align: center; font-size: 11px;
    border-right: 1px solid var(--border-strong); border-bottom: 1px solid var(--gridline);
    font-family: Calibri, "Segoe UI", sans-serif; line-height: 21px;
  }
  .hx-cell {
    flex: 0 0 auto; padding: 1px 5px; background: #fff;
    border-right: 1px solid var(--gridline); border-bottom: 1px solid var(--gridline);
    font-family: Calibri, "Microsoft YaHei", sans-serif; font-size: 13px; color: var(--text);
    line-height: 19px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
  }
  .hx-cell.num { text-align: right; font-variant-numeric: tabular-nums; }
  .hx-cell.wrap { white-space: normal; word-break: break-word; }
  .hx-cell.content { white-space: normal; word-break: break-word; max-height: 168px; overflow: auto; padding: 3px 6px; line-height: 1.5; }
  .hx-cell.content::-webkit-scrollbar { width: 8px; height: 8px; }
  .hx-cell.content::-webkit-scrollbar-thumb { background: #c9c9c9; border-radius: 4px; }
  .hx-cell.sel { outline: 2px solid var(--accent); outline-offset: -2px; position: relative; z-index: 1; }
  .hx-coll.sel, .hx-rowhead.sel { background: var(--accent-soft); color: var(--accent-dark); font-weight: 700; }
  .hx-link { color: var(--link); text-decoration: underline; }
  .hx-link:hover { color: var(--accent-dark); }

  /* 分组标题行（板块分区那种「分类 + 该分类下的板块」的层级表） */
  .hx-group-cell {
    background: #eef4f1; color: #175c38; font-weight: 600; font-size: 12px;
    border-right: 1px solid var(--gridline); border-bottom: 1px solid #d5e4dd;
    line-height: 20px; padding: 2px 8px;
  }
  .hx-group-cell .hx-link { color: #175c38; text-decoration: none; }
  .hx-group-cell .hx-link:hover { text-decoration: underline; }
  .hx-group-note { color: #7c8f87; font-weight: 400; font-size: 11px; }
  .hx-tr-group .hx-rowhead { background: #eef4f1; color: #175c38; }
  #hx-root.hx-nofreeze .hx-tr.hx-th, #hx-root.hx-nofreeze .hx-rowhead, #hx-root.hx-nofreeze .hx-corner { position: static; }

  /* 账号区 */
  .hx-account { text-decoration: none; }
  .hx-account .hx-uname { font-weight: 600; max-width: 120px; overflow: hidden; text-overflow: ellipsis; }
  .hx-avatar-img { width: 20px; height: 20px; border-radius: 50%; object-fit: cover; display: block; }

  /* Excel 选项对话框 */
  .hx-dlg-wrap { position: absolute; inset: 0; z-index: 20; }
  .hx-dlg-mask { position: absolute; inset: 0; background: rgba(0,0,0,.18); }
  .hx-dlg {
    position: absolute; left: 50%; top: 46px; transform: translateX(-50%);
    width: 780px; max-width: calc(100% - 40px); height: 540px; max-height: calc(100% - 100px);
    background: #fff; border: 1px solid #c8c8c8; box-shadow: 0 10px 30px rgba(0,0,0,.24);
    display: flex; flex-direction: column; font-size: 12px; color: #201f1e;
  }
  .hx-dlg-title {
    flex: 0 0 auto; height: 34px; display: flex; align-items: center; padding: 0 6px 0 12px;
    border-bottom: 1px solid var(--border); background: var(--head-bg); font-weight: 600;
  }
  .hx-dlg-x { margin-left: auto; width: 32px; height: 26px; display: flex; align-items: center; justify-content: center; font-style: normal; border-radius: 3px; cursor: default; }
  .hx-dlg-x:hover { background: #e81123; color: #fff; }
  .hx-dlg-body { flex: 1 1 auto; display: flex; min-height: 0; }
  .hx-dlg-nav { flex: 0 0 150px; background: #faf9f8; border-right: 1px solid var(--border); padding: 8px 6px; }
  .hx-dlg-navitem { padding: 7px 10px; border-radius: 3px; color: #333; cursor: default; }
  .hx-dlg-navitem:hover { background: #f0efee; }
  .hx-dlg-navitem.active { background: var(--accent-soft); color: var(--accent-dark); font-weight: 600; }
  .hx-dlg-main { flex: 1 1 auto; padding: 12px 18px; overflow: auto; }
  .hx-dlg-main::-webkit-scrollbar { width: 10px; }
  .hx-dlg-main::-webkit-scrollbar-thumb { background: #c9c9c9; border-radius: 5px; }
  .hx-set-group { display: none; }
  .hx-set-group.active { display: block; }
  .hx-set-group-title { font-size: 13px; font-weight: 600; color: var(--accent-dark); margin: 2px 0 8px; }
  .hx-set-row { display: flex; align-items: center; gap: 14px; padding: 10px 2px; border-bottom: 1px solid #f0f0f0; }
  .hx-set-main { flex: 1 1 auto; min-width: 0; }
  .hx-set-title { font-size: 12px; font-weight: 600; color: #201f1e; }
  .hx-set-desc { font-size: 11px; color: #8a8886; margin-top: 3px; line-height: 1.5; }
  .hx-switch { flex: 0 0 auto; position: relative; width: 40px; height: 20px; }
  .hx-switch input { position: absolute; inset: 0; width: 100%; height: 100%; opacity: 0; margin: 0; cursor: pointer; }
  .hx-switch span { position: absolute; inset: 0; background: #c8c6c4; border-radius: 10px; transition: background .15s; pointer-events: none; }
  .hx-switch span::after { content: ""; position: absolute; top: 3px; left: 3px; width: 14px; height: 14px; border-radius: 50%; background: #fff; transition: transform .15s; }
  .hx-switch input:checked + span { background: var(--accent); }
  .hx-switch input:checked + span::after { transform: translateX(20px); }
  .hx-dlg select, .hx-dlg input[type="text"] {
    flex: 0 0 auto; min-width: 190px; height: 26px; padding: 0 6px; border: 1px solid var(--border-strong);
    border-radius: 2px; background: #fff; font: inherit; color: #201f1e; outline: none;
  }
  .hx-dlg select:focus, .hx-dlg input[type="text"]:focus { border-color: var(--accent); }
  /* 滑块（缩略图尺寸） */
  .hx-range { flex: 0 0 auto; display: flex; align-items: center; gap: 12px; }
  .hx-range input[type="range"] {
    -webkit-appearance: none; appearance: none; width: 220px; height: 4px; border-radius: 2px;
    background: linear-gradient(to right, var(--accent) 0, var(--accent) var(--fill, 50%), #dcdcdc var(--fill, 50%), #dcdcdc 100%);
    outline: none; cursor: pointer;
  }
  .hx-range input[type="range"]::-webkit-slider-thumb {
    -webkit-appearance: none; appearance: none; width: 15px; height: 15px; border-radius: 50%;
    background: var(--accent); border: 2px solid #fff; box-shadow: 0 1px 4px rgba(0,0,0,.3); cursor: pointer;
  }
  .hx-range input[type="range"]::-moz-range-thumb {
    width: 13px; height: 13px; border-radius: 50%; background: var(--accent);
    border: 2px solid #fff; box-shadow: 0 1px 4px rgba(0,0,0,.3); cursor: pointer;
  }
  .hx-range-val { flex: 0 0 54px; text-align: right; font-size: 11px; color: #605e5c; font-variant-numeric: tabular-nums; }
  .hx-key-row { display: flex; align-items: center; gap: 12px; padding: 8px 2px; border-bottom: 1px solid #f0f0f0; }
  .hx-key-row kbd {
    flex: 0 0 168px; font-family: Consolas, "Courier New", monospace; font-size: 11px; color: #201f1e;
    background: #f3f2f1; border: 1px solid #e1dfdd; border-bottom-width: 2px; border-radius: 3px; padding: 3px 6px; text-align: center;
  }
  .hx-dlg-foot { flex: 0 0 auto; height: 46px; display: flex; align-items: center; gap: 10px; padding: 0 14px; border-top: 1px solid var(--border); background: #faf9f8; }
  .hx-dlg-btn { height: 28px; min-width: 96px; padding: 0 14px; border: 1px solid var(--border-strong); background: #fff; border-radius: 2px; font: inherit; color: #201f1e; cursor: pointer; }
  .hx-dlg-btn:hover { background: #f3f2f1; }
  .hx-dlg-btn.primary { background: var(--accent); border-color: var(--accent); color: #fff; }
  .hx-dlg-btn.primary:hover { filter: brightness(1.08); }
  .hx-dlg-tip { margin-left: auto; font-size: 11px; color: #8a8886; }
  .hx-quote { border-left: 3px solid #c8c8c8; background: #fafafa; padding: 2px 8px; margin: 3px 0; color: #6b6b6b; }
  .hx-img { display: inline-block; max-width: var(--img-max-w); max-height: var(--img-max-h); margin: 3px 4px 3px 0; border: 1px solid #e0e0e0; background: #fafafa; vertical-align: top; cursor: zoom-in; }

  /* 鼠标悬停浮出大图 */
  .hx-zoom-pop {
    position: fixed; left: -9999px; top: -9999px; z-index: 40; display: none;
    min-width: 90px; min-height: 64px;
    padding: 4px; background: #fff; border: 1px solid #bfbfbf; border-radius: 3px;
    box-shadow: 0 8px 28px rgba(0,0,0,.28); pointer-events: none;
  }
  .hx-zoom-pop.on { display: block; }
  .hx-zoom-pop img { display: block; max-width: 62vw; max-height: 72vh; background: #fafafa; }
  .hx-zoom-pop .hx-zoom-cap {
    display: block; max-width: 62vw; padding: 4px 2px 1px; font-size: 11px; color: #8a8886;
    overflow: hidden; white-space: nowrap; text-overflow: ellipsis;
  }
  #hx-root.hx-nozoom .hx-img { cursor: default; }
  .hx-empty { padding: 14px 12px; color: var(--muted); font-size: 13px; }
  .hx-loading { padding: 24px 16px; color: var(--muted); font-size: 13px; }
  .hx-loading b { color: var(--accent); font-weight: 600; }

  /* ---------- 底部：工作表标签 / 分页 / 状态栏 ---------- */
  .hx-status {
    flex: 0 0 auto; height: 28px; display: flex; align-items: stretch;
    background: #fff; border-top: 1px solid var(--border); font-size: 11px; color: #444;
  }
  .hx-nav4 { display: flex; align-items: center; gap: 1px; padding: 0 6px; color: #8a8886; }
  .hx-nav4 i { width: 20px; height: 22px; display: flex; align-items: center; justify-content: center; font-style: normal; }
  .hx-nav4 i:hover { background: rgba(0,0,0,.06); }
  .hx-tabstrip { display: flex; align-items: stretch; overflow-x: auto; overflow-y: hidden; }
  .hx-tabstrip::-webkit-scrollbar { height: 0; }
  .hx-sheet-tab {
    display: flex; align-items: center; padding: 0 14px; white-space: nowrap; cursor: default;
    background: #f3f2f1; color: #444; border-right: 1px solid var(--border); position: relative; max-width: 220px;
  }
  .hx-sheet-tab span { overflow: hidden; text-overflow: ellipsis; }
  .hx-sheet-tab:hover { background: #eaeaea; }
  .hx-sheet-tab.active { background: #fff; color: var(--accent); font-weight: 600; }
  .hx-sheet-tab.active::after { content: ""; position: absolute; left: 0; right: 0; bottom: 0; height: 2px; background: var(--accent); }
  .hx-sheet-add { display: flex; align-items: center; padding: 0 9px; color: #8a8886; border-right: 1px solid var(--border); }
  .hx-sheet-add:hover { background: rgba(0,0,0,.06); }
  .hx-pager { margin-left: auto; display: flex; align-items: center; gap: 8px; padding: 0 10px; border-left: 1px solid var(--border); color: var(--muted); white-space: nowrap; }
  .hx-pager b { color: #333; font-weight: 600; }
  .hx-pager a, .hx-pager button { color: var(--link); background: none; border: none; font: inherit; cursor: pointer; padding: 0 2px; }
  .hx-pager a.off { color: #bbb; pointer-events: none; }
  .hx-pager input {
    width: 34px; height: 18px; border: 1px solid var(--border-strong); border-radius: 2px;
    text-align: center; font: inherit; color: #333; background: #fff; outline: none;
  }
  .hx-status-right { display: flex; align-items: center; gap: 8px; padding: 0 10px; border-left: 1px solid var(--border); color: var(--muted); }
  .hx-status-right svg { width: 14px; height: 14px; }
  .hx-zoom { display: flex; align-items: center; gap: 6px; }
  .hx-zoom i { font-style: normal; font-size: 14px; color: #666; }
  .hx-zoombar { width: 74px; height: 3px; background: #d6d6d6; border-radius: 2px; position: relative; }
  .hx-zoombar::after { content: ""; position: absolute; left: 62%; top: -4px; width: 10px; height: 10px; border-radius: 50%; background: #fff; border: 1px solid #9a9a9a; }
  .hx-hint {
    position: fixed; right: 14px; bottom: 44px; z-index: 2147483001;
    background: rgba(32,31,30,.92); color: #fff; font-size: 12px; padding: 7px 12px; border-radius: 4px;
    opacity: 0; transition: opacity .18s ease; pointer-events: none;
  }
  .hx-hint.show { opacity: 1; }
  `;

  /**
   * 给所有选择器加上 #hx-root 前缀。
   * 因为 scope reset 用的是 `#hx-root *`（1 个 ID + 通配），特异性是 (1,0,0)，
   * 而组件规则如果只有类选择器就是 (0,1,0) —— ID 会赢，reset 会把组件的
   * border/padding/max-width 全踩掉。统一加上 #hx-root 前缀后两边都是 ID 级，
   * 由后面的类选择器正常覆盖通配选择器。
   */
  function scopeCss(css) {
    return css.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|\})\s*([^{}@]+)\{/g, (m, pre, sel) => {
      const scoped = sel.split(',').map(s => {
        s = s.trim();
        if (!s) return s;
        if (/^(#hx-root|html\.)/.test(s)) return s;
        return '#hx-root ' + s;
      }).join(', ');
      return pre + scoped + '{';
    });
  }

  /*
   * 首帧就要生效的三行单独抽出来，注入时不经过 scopeCss 的正则加工 ——
   * 从「脚本开始跑」到「浏览器算完样式准备画第一帧」之间只有几毫秒，
   * 这里省掉的是 16KB 正则替换的时间。规则本身已经是全限定的，直接塞就行。
   */
  const HIDE_CSS =
    'html.hx-on,html.hx-on body{overflow:hidden!important;background:#fff!important}' +
    'html.hx-on body > *:not(#hx-root),html.hx-on > *:not(head):not(body):not(#hx-root){display:none!important}' +
    'html.hx-peek #hx-root{display:none!important}';

  function injectCss() {
    if (!document.getElementById('hx-hide')) {
      const hide = document.createElement('style');
      hide.id = 'hx-hide';
      hide.textContent = HIDE_CSS;
      const parent = document.head || document.documentElement;
      if (parent) parent.appendChild(hide);
    }
    let style = document.getElementById('hx-style');
    if (!style) {
      style = document.createElement('style');
      style.id = 'hx-style';
      style.textContent = scopeCss(CSS);
      const parent = document.head || document.documentElement;
      if (parent) parent.appendChild(style);
      return;
    }
    // 早期注入时 <head> 可能还没被解析出来，DOM 就绪后挪进去
    if (document.head && style.parentNode !== document.head) document.head.appendChild(style);
  }

  /**
   * document-start 阶段 document.documentElement 可能还不存在
   * （Chrome 在 run_at:document_start 时给的就是一个空 document），
   * 所有依赖 html 元素的操作都要先等它出现。
   */
  function whenDocumentElement(fn) {
    if (document.documentElement) { fn(); return; }
    const obs = new MutationObserver(() => {
      if (!document.documentElement) return;
      obs.disconnect();
      fn();
    });
    obs.observe(document, { childList: true });
  }

  /** 同理，等 <body> 出现 */
  function whenBody(fn) {
    if (document.body) { fn(); return; }
    const target = document.documentElement || document;
    const obs = new MutationObserver(() => {
      if (!document.body) return;
      obs.disconnect();
      fn();
    });
    obs.observe(target, { childList: true, subtree: true });
  }

  /* ============================== 5. 图标（内联 SVG，无外部依赖） ============================== */
  /*
   * nga-excel.js 的 Excel 皮肤是从 NGA 的资源脚本里取 base64 图片的，
   * 这里不依赖任何外部资源，全部用小尺寸线性 SVG + 文字标签手绘，
   * 配合功能区每个按钮下方的中文说明，缩略看和真 Excel 很像。
   */
  const SVG_OPEN = '<svg viewBox="0 0 20 20">';
  const ICONS = {
    save: SVG_OPEN + '<path d="M4 3h9l3 3v11H4z"/><path d="M7 3v5h6V3"/><rect x="6.5" y="11" width="7" height="6"/></svg>',
    undo: SVG_OPEN + '<path d="M4 11h8.5a4 4 0 0 1 0 8H8"/><path d="M7.5 7 4 11l3.5 4"/></svg>',
    redo: SVG_OPEN + '<path d="M16 11H7.5a4 4 0 0 0 0 8H12"/><path d="M12.5 7 16 11l-3.5 4"/></svg>',
    share: SVG_OPEN + '<path d="M10 3v9"/><path d="M6.5 6.5 10 3l3.5 3.5"/><path d="M4.5 12v4.5h11V12"/></svg>',
    comment: SVG_OPEN + '<path d="M4 4h12v9H9.5L5.5 16v-3H4z"/></svg>',
    bell: SVG_OPEN + '<path d="M10 3a4 4 0 0 1 4 4v3.2L15.5 13H4.5L6 10.2V7a4 4 0 0 1 4-4z"/><path d="M8.5 15.5a1.6 1.6 0 0 0 3 0"/></svg>',
    user: SVG_OPEN + '<circle cx="10" cy="7.2" r="3"/><path d="M4.5 17c.3-3 2.6-4.8 5.5-4.8S15.2 14 15.5 17"/></svg>',
    search: SVG_OPEN + '<circle cx="9" cy="9" r="5.5"/><path d="m13.2 13.2 3.8 3.8"/></svg>',
    chevron: SVG_OPEN + '<path d="m6 8 4 4 4-4"/></svg>',
    caretDown: SVG_OPEN + '<path d="m5 5 5 5 5-5z" fill="currentColor" stroke="none"/><path d="m5 11 5 5 5-5z" fill="currentColor" stroke="none" opacity=".45"/></svg>',
    caretUp: SVG_OPEN + '<path d="m5 15 5-5 5 5z" fill="currentColor" stroke="none"/><path d="m5 9 5-5 5 5z" fill="currentColor" stroke="none" opacity=".45"/></svg>',

    paste: SVG_OPEN + '<rect x="4" y="4" width="12" height="14" rx="1"/><rect x="7" y="2" width="6" height="4" rx="1"/><path d="M7 10.5h6M7 13.5h4"/></svg>',
    scissors: SVG_OPEN + '<circle cx="5.5" cy="15" r="2"/><circle cx="5.5" cy="5" r="2"/><path d="M7.3 13.7 16 4M7.3 6.3 16 16"/></svg>',
    copy: SVG_OPEN + '<rect x="6.5" y="6.5" width="10" height="11" rx="1"/><path d="M13.5 6.5V4a1 1 0 0 0-1-1h-8a1 1 0 0 0-1 1v9a1 1 0 0 0 1 1h2"/></svg>',
    brush: SVG_OPEN + '<path d="M4 16.5c2.4 0 3.6-1.3 3.6-3.2"/><path d="M7.2 13 14.4 5.2 16.8 7.6 9.6 15.4z"/><path d="m14.4 5.2 1.4-1.4"/></svg>',
    fill: SVG_OPEN + '<path d="M6.6 3.5 13 9.9l-5 5-4.6-4.6z"/><path d="M9 6 11.6 3.4"/><path d="M3.5 17.5h13" stroke-width="2.4"/></svg>',
    border: SVG_OPEN + '<rect x="3" y="3" width="14" height="14"/><path d="M10 3v14M3 10h14"/></svg>',
    alignLeft: SVG_OPEN + '<path d="M3 4.5h14M3 8.5h9M3 12.5h14M3 16.5h9"/></svg>',
    alignCenter: SVG_OPEN + '<path d="M3 4.5h14M5.5 8.5h9M3 12.5h14M5.5 16.5h9"/></svg>',
    alignRight: SVG_OPEN + '<path d="M3 4.5h14M8 8.5h9M3 12.5h14M8 16.5h9"/></svg>',
    alignTop: SVG_OPEN + '<path d="M4.5 3v14M8.5 3v9M12.5 3v14M16.5 3v9"/></svg>',
    alignMiddle: SVG_OPEN + '<path d="M4.5 3v14M8.5 5.5v9M12.5 3v14M16.5 5.5v9"/></svg>',
    alignBottom: SVG_OPEN + '<path d="M4.5 3v14M8.5 8v9M12.5 3v14M16.5 8v9"/></svg>',
    indentDec: SVG_OPEN + '<path d="M9 4.5h8M9 10h8M9 15.5h8M6.5 7 3.5 10l3 3z"/></svg>',
    indentInc: SVG_OPEN + '<path d="M9 4.5h8M9 10h8M9 15.5h8M3.5 7l3 3-3 3z"/></svg>',
    merge: SVG_OPEN + '<rect x="2.5" y="5.5" width="15" height="9" rx="1"/><path d="M10 8v4M8 10.5 10 12.5l2-2"/></svg>',
    wrap: SVG_OPEN + '<path d="M3 5h14M3 9.5h9a2.5 2.5 0 0 1 0 5H9"/><path d="M10.5 12.5 9 14.5l1.5 2"/></svg>',
    condFmt: SVG_OPEN + '<rect x="3" y="4" width="14" height="12"/><path d="M3 8h14M8 8v8"/><rect x="3" y="4" width="5" height="4" fill="currentColor" stroke="none" opacity=".35"/><rect x="3" y="12" width="5" height="4" fill="currentColor" stroke="none" opacity=".18"/></svg>',
    tableFmt: SVG_OPEN + '<rect x="3" y="4" width="14" height="12"/><path d="M3 7h14M3 10h14M3 13h14M8 7v9M13 7v9"/><rect x="3" y="4" width="14" height="3" fill="currentColor" stroke="none" opacity=".35"/></svg>',
    cellStyle: SVG_OPEN + '<rect x="3" y="4" width="14" height="12"/><path d="M3 8h14M3 12h14M8 4v12M13 4v12"/></svg>',
    insertCell: SVG_OPEN + '<rect x="3.5" y="3.5" width="13" height="13" rx="1"/><path d="M10 7v6M7 10h6"/></svg>',
    deleteCell: SVG_OPEN + '<path d="M5 5.5h10M8.2 5.5V3.8h3.6v1.7"/><path d="M6.6 5.5l.9 11.5h5l.9-11.5"/><path d="M9 8.5v6M11 8.5v6"/></svg>',
    formatCell: SVG_OPEN + '<path d="M5.5 17 10 3.5 14.5 17"/><path d="M7.2 12.6h5.6"/><path d="M3 17h14"/></svg>',
    sum: SVG_OPEN + '<path d="M5.5 4.5h9l-5 5.5 5 5.5h-9"/></svg>',
    fillDown: SVG_OPEN + '<rect x="4" y="3" width="12" height="4.5" rx="1"/><path d="M10 10v6"/><path d="M7.5 13.5 10 16l2.5-2.5"/></svg>',
    clear: SVG_OPEN + '<path d="M4 17h12"/><path d="m6 15 8-8 3 3-5.2 5z"/><path d="m14 7 2.2-2.2"/></svg>',
    sort: SVG_OPEN + '<path d="M5 3v13"/><path d="M3 14l2 3 2-3"/><path d="M10 5.5h7M10 9.5h5M10 13.5h3"/></svg>',
    find: SVG_OPEN + '<circle cx="9" cy="9" r="5.5"/><path d="m13.2 13.2 3.8 3.8"/><path d="M6.5 9h5"/></svg>',
    viewNormal: SVG_OPEN + '<rect x="3" y="4.5" width="14" height="11"/><path d="M6 8h8M6 11h5"/></svg>',
    viewLayout: SVG_OPEN + '<rect x="3" y="4.5" width="14" height="11"/><path d="M3 8h14M10 8v7.5"/></svg>',
    viewPage: SVG_OPEN + '<rect x="4.5" y="3" width="11" height="14"/><path d="M7 7h6M7 10h6M7 13h3"/></svg>',
    gear: SVG_OPEN + '<circle cx="10" cy="10" r="3"/><path d="M10 2v2.4M10 15.6V18M2 10h2.4M15.6 10H18M4.4 4.4l1.7 1.7M13.9 13.9l1.7 1.7M15.6 4.4l-1.7 1.7M6.1 13.9l-1.7 1.7"/></svg>',
    home: SVG_OPEN + '<path d="M3 9.4 10 3.2l7 6.2"/><path d="M5.2 8.6V16.8h9.6V8.6"/><path d="M8.3 16.8v-4.2h3.4v4.2"/></svg>'
  };

  const glyph = (t, cls, extra) =>
    '<span class="hx-glyph ' + cls + '"' + (extra || '') + '>' + t + '</span>';

  /** 功能区按钮：图标 + 中文小字（真 Excel 也是这么排的） */
  function gbtn(icon, label, opts) {
    const o = opts || {};
    const inner = ICONS[icon] || (/^</.test(icon) ? icon : '');
    return '<div class="hx-btn' + (o.cls ? ' ' + o.cls : '') + '"' +
      (o.title ? ' title="' + o.title + '"' : '') + '>' +
      '<div class="i">' + inner + '</div>' +
      (label ? '<div class="t">' + label + '</div>' : '') + '</div>';
  }

  function fontGroup() {
    return '<div class="hx-gbody" style="flex-direction:column;align-items:stretch;gap:3px;justify-content:center">' +
      '<div style="display:flex;align-items:center;gap:4px">' +
      '<div class="hx-combo wide"><span>等线 (中文正文)</span>' + ICONS.chevron + '</div>' +
      '<div class="hx-combo" style="min-width:42px"><span>11</span>' + ICONS.chevron + '</div>' +
      gbtn(ICONS.caretUp, '', { cls: 'hx-tiny' }) + gbtn(ICONS.caretDown, '', { cls: 'hx-tiny' }) +
      '</div>' +
      '<div style="display:flex;align-items:center;gap:1px">' +
      gbtn(glyph('B', 'b'), '') + gbtn(glyph('I', 'i'), '') + gbtn(glyph('U', 'u'), '') + gbtn(glyph('S', 's'), '') +
      '<div style="width:1px;height:18px;background:#e3e3e3;margin:0 3px"></div>' +
      gbtn(glyph('A') + '<span class="hx-bar" style="background:#c0392b"></span>', '字体颜色') +
      gbtn(ICONS.fill + '<span class="hx-bar" style="background:#ffd966"></span>', '填充颜色') +
      gbtn(ICONS.border, '边框') +
      '</div></div>';
  }

  function alignGroup() {
    return '<div class="hx-gbody" style="flex-direction:column;align-items:stretch;gap:1px;justify-content:center">' +
      '<div style="display:flex;gap:1px">' +
      gbtn(ICONS.alignTop, '顶端对齐') + gbtn(ICONS.alignMiddle, '垂直居中') + gbtn(ICONS.alignBottom, '底端对齐') +
      '</div>' +
      '<div style="display:flex;gap:1px;align-items:center">' +
      gbtn(ICONS.alignLeft, '左对齐') + gbtn(ICONS.alignCenter, '居中') + gbtn(ICONS.alignRight, '右对齐') +
      gbtn(ICONS.indentDec, '') + gbtn(ICONS.indentInc, '') +
      '<div style="width:1px;height:18px;background:#e3e3e3;margin:0 3px"></div>' +
      gbtn(ICONS.merge, '合并后居中') + gbtn(ICONS.wrap, '自动换行') +
      '</div></div>';
  }

  function numberGroup() {
    // 和真 Excel 一样：一个「常规」下拉 + 5 个无文字的格式按钮
    return '<div class="hx-numrow">' +
      '<div class="hx-combo" style="min-width:64px"><span>常规</span>' + ICONS.chevron + '</div>' +
      gbtn(glyph('￥'), '', { title: '会计专用格式' }) +
      gbtn(glyph('%'), '', { title: '百分比样式' }) +
      gbtn(glyph('，'), '', { title: '千位分隔样式' }) +
      gbtn(glyph('.00'), '', { title: '增加小数位数' }) +
      gbtn(glyph('.0'), '', { title: '减少小数位数' }) +
      '</div>';
  }

  function styleGroup() {
    return '<div class="hx-gbody" style="align-items:center">' +
      gbtn(ICONS.condFmt, '条件格式') + gbtn(ICONS.tableFmt, '表格格式') + gbtn(ICONS.cellStyle, '单元格样式') +
      '</div>';
  }

  const RIBBON = [
    { title: '剪贴板', html: gbtn('paste', '粘贴', { big: true }) + gbtn('scissors', '剪切') + gbtn('copy', '复制') + gbtn('brush', '格式刷') },
    { title: '字体', html: fontGroup() },
    { title: '对齐方式', html: alignGroup() },
    { title: '数字', html: numberGroup() },
    { title: '样式', html: styleGroup() },
    { title: '单元格', html: gbtn('insertCell', '插入') + gbtn('deleteCell', '删除') + gbtn('formatCell', '格式') },
    { title: '编辑', html: gbtn('sum', '求和') + gbtn('fillDown', '填充') + gbtn('clear', '清除') + gbtn('sort', '排序') + gbtn('find', '查找') }
  ];

  const TAB_NAMES = ['文件', '开始', '插入', '页面布局', '公式', '数据', '审阅', '视图', '安全', '开发工具', '特色功能'];

  function ribbonHtml() {
    return RIBBON.map(g =>
      '<div class="hx-group"><div class="hx-gbody">' + g.html + '</div><div class="hx-gtitle">' + g.title + '</div></div>'
    ).join('');
  }

  function excelLogoSvg() {
    return '<svg viewBox="0 0 32 32"><rect x="3" y="3" width="26" height="26" rx="4" fill="#217346" stroke="none"/>' +
      '<path d="M20.6 10.4 12 21.6M12 10.4l8.6 11.2" stroke="#fff" stroke-width="2.6" fill="none" stroke-linecap="round"/></svg>';
  }

  /* ============================== 6. 界面骨架 ============================== */

  const state = {
    model: null,
    sheet: 0,
    sel: { r: 0, c: 0 },
    cells: [],
    tabs: [],
    colHeads: []
  };
  let R = null;

  function chromeHtml() {
    const tabs = TAB_NAMES.map((t, i) =>
      '<div class="hx-tab' + (t === '文件' ? ' file' : (t === '开始' ? ' active' : '')) + '">' + t + '</div>'
    ).join('');

    return '' +
      '<div class="hx-titlebar">' +
        '<div class="hx-logo">' + excelLogoSvg() + '</div>' +
        '<div class="hx-book">工作簿1</div>' +
        '<div class="hx-save"><span class="hx-dot"></span><span>已保存到此电脑</span></div>' +
        '<div class="hx-tbtn">' + ICONS.save + '</div>' +
        '<div class="hx-tbtn">' + ICONS.undo + '</div>' +
        '<div class="hx-tbtn">' + ICONS.redo + '</div>' +
        '<div class="hx-spacer"></div>' +
        '<input class="hx-search" type="text" placeholder="搜索（在虎扑站内搜索）" spellcheck="false">' +
        '<div class="hx-acct"></div>' +
        '<div class="hx-tbtn hx-home" title="返回社区首页">' + ICONS.home + '</div>' +
        '<div class="hx-tbtn hx-gear" title="设置（Excel 选项）">' + ICONS.gear + '</div>' +
        '<div class="hx-win"><i>−</i><i>▢</i><i class="close">✕</i></div>' +
      '</div>' +
      '<div class="hx-tabs">' + tabs +
        '<div class="hx-tabs-right">' + ICONS.search + '<span>查找</span></div>' +
      '</div>' +
      '<div class="hx-ribbon">' + ribbonHtml() + '</div>' +
      '<div class="hx-formulabar">' +
        '<div class="hx-namebox">A1</div>' +
        '<div class="hx-fx">fx</div>' +
        '<div class="hx-formula"></div>' +
      '</div>' +
      '<div class="hx-sheet">' +
        '<div class="hx-table"><div class="hx-loading">正在打开 <b>' + esc(CFG.book || '工作簿1') + '.xlsx</b> …</div></div>' +
      '</div>' +
      '<div class="hx-status">' +
        '<div class="hx-nav4"><i>|◀</i><i>◀</i><i>▶</i><i>▶|</i></div>' +
        '<div class="hx-tabstrip"></div>' +
        '<div class="hx-sheet-add">+</div>' +
        '<div class="hx-pager"></div>' +
        '<div class="hx-status-right">' +
          '<span class="hx-vbtn">' + ICONS.viewNormal + '</span>' +
          '<span class="hx-vbtn">' + ICONS.viewLayout + '</span>' +
          '<span class="hx-vbtn">' + ICONS.viewPage + '</span>' +
          '<span class="hx-zoom"><i>−</i><span class="hx-zoombar"></span><i>+</i></div>' +
          '<span>100%</span>' +
        '</div>' +
      '</div>' +
      '<div class="hx-hint"></div>';
  }

  function ensureRoot() {
    if (R && document.body && document.body.contains(R.root)) return R;
    const root = el('div');
    root.id = 'hx-root';
    root.className = rootClass();
    root.innerHTML = chromeHtml();
    (document.body || document.documentElement).appendChild(root);

    R = {
      root: root,
      book: root.querySelector('.hx-book'),
      acct: root.querySelector('.hx-acct'),
      search: root.querySelector('.hx-search'),
      formula: root.querySelector('.hx-formula'),
      namebox: root.querySelector('.hx-namebox'),
      table: root.querySelector('.hx-table'),
      tabstrip: root.querySelector('.hx-tabstrip'),
      pager: root.querySelector('.hx-pager'),
      hint: root.querySelector('.hx-hint')
    };
    bindRoot();
    return R;
  }

  function toast(msg, ms) {
    if (!R) return;
    R.hint.textContent = msg;
    R.hint.classList.add('show');
    clearTimeout(toast._t);
    toast._t = setTimeout(() => R.hint.classList.remove('show'), ms || 2600);
  }

  /* ============================== 7. 渲染 ============================== */

  function fallbackModel() {
    return {
      title: pageTitleText() || document.title,
      crumbs: crumbs(),
      crumbList: crumbList(),
      sheetName: 'Sheet1',
      pager: null,
      sheets: [{
        name: 'Sheet1',
        cols: [{ label: '说明', width: 900, wrap: true }],
        rows: [[{ text: '这个页面暂时没有可以摊成表格的数据（脚本主要接管 bbs.hupu.com 的版面页 / 帖子页）。按 Esc 切回虎扑原页面，Alt+E 关闭 Excel 模式。' }]]
      }]
    };
  }

  /**
   * @param keepSheet 同一页面的重绘（改设置、看门狗补渲染）时传 true：尽量留在
   *   用户原来那张工作表上，别把人踹回第一张。页面切换（软导航）不传：那是进了
   *   新页面，要落到新页面自己的默认工作表（sheetName / 第一张）。
   *
   *   以前整页跳转时有「版面 A → 版面 B 自动显示 B」的效果，正是因为新文档
   *   state.model 是空的、只能从 0 开始；软导航不会重置 state，所以得在这儿显式区分。
   */
  function paint(model, keepSheet) {
    const prevSheet = state.model && state.model.sheets[state.sheet] ? state.model.sheets[state.sheet].name : null;
    state.model = model;
    state.sheet = 0;
    state.sel = { r: 0, c: 0 };
    if (keepSheet && prevSheet) {
      for (let i = 1; i < model.sheets.length; i++) {
        if (model.sheets[i].name === prevSheet) { state.sheet = i; break; }
      }
    } else if (model.sheetName) {
      // 新页面的默认工作表（listFromData / threadFromNext / modelHome 都会给）
      for (let i = 0; i < model.sheets.length; i++) {
        if (model.sheets[i].name === model.sheetName) { state.sheet = i; break; }
      }
    }
    R.root.className = rootClass();
    applyTweaks();

    const book = CFG.book || model.title || '工作簿1';
    wantTitle = book;
    R.book.textContent = book;

    renderAccount();
    // 进新页面时表格从头看起（和整页跳转一致）；同一页面重绘则保留滚动位置
    if (!keepSheet && R.table && R.table.parentNode) {
      R.table.parentNode.scrollTop = 0;
      R.table.parentNode.scrollLeft = 0;
    }
    renderSheet();
  }

  /** #hx-root 的类名：皮肤 + 冻结开关 */
  function rootClass() {
    return 'hx-t-' + CFG.theme + (CFG.freezeHeader ? '' : ' hx-nofreeze');
  }

  /**
   * 不用重建表格就能生效的配置：图片尺寸（CSS 变量）与悬停放大开关。
   * 设置面板里拖动滑块时只走这里，不会重建整个表。
   */
  function applyTweaks() {
    if (!R) return;
    R.root.style.setProperty('--img-max-w', clamp(num(CFG.imgMaxW) || 260, 60, 1200) + 'px');
    R.root.style.setProperty('--img-max-h', clamp(num(CFG.imgMaxH) || 170, 40, 900) + 'px');
    R.root.classList.toggle('hx-nozoom', !CFG.imgHoverZoom);
    if (!CFG.imgHoverZoom) hideZoomPop();
  }

  /**
   * 编辑栏固定显示面包屑（可点的链接）» 标题。
   *
   * 它**不会**因为选中单元格而被单元格内容顶掉（以前会，结果点一下回复就找不到
   * 「回上一层」的入口了）。单元格内容本来就在格子里显示，这一行留给导航和标题。
   */
  function showTrail() {
    const model = state.model;
    if (!model || !R) return;
    const list = (model.crumbList || []).slice();
    if (!list.length && model.crumbs) {
      model.crumbs.forEach(t => list.push({ title: t, url: '' }));
    }
    // 面包屑最后一节和页面标题重复时只留一个
    if (model.title && list.length && list[list.length - 1].title === model.title) list.pop();

    R.formula.innerHTML = '';
    list.forEach(c => {
      if (c.url) {
        const a = el('a', 'hx-link');
        a.href = c.url;
        a.textContent = c.title;
        R.formula.appendChild(a);
      } else {
        R.formula.appendChild(document.createTextNode(c.title));
      }
      R.formula.appendChild(document.createTextNode(' » '));
    });
    R.formula.appendChild(document.createTextNode(model.title || ''));
  }

  /** 按配置过滤列（「路径」列可以关掉，默认关） */
  function visibleCols(sheet) {
    const all = sheet.cols || [];
    const out = [];
    all.forEach((col, idx) => {
      if (col.url && !CFG.showUrl) return;
      out.push({ col: col, idx: idx });
    });
    return out.length ? out : all.map((col, idx) => ({ col: col, idx: idx }));
  }

  function renderSheet() {
    const model = state.model;
    const sheet = model.sheets[state.sheet] || model.sheets[0];
    const table = R.table;
    table.innerHTML = '';
    state.cells = [];
    state.colHeads = [];

    if (!sheet || !sheet.rows || !sheet.rows.length) {
      table.appendChild(el('div', 'hx-empty', '这张工作表里没有数据。'));
      renderTabs();
      showTrail();
      return;
    }

    // 列标行
    const head = el('div', 'hx-tr hx-th');
    const corner = el('div', 'hx-corner');
    head.appendChild(corner);
    const cols = visibleCols(sheet);
    const tableWidth = cols.reduce((s, c) => s + c.col.width, 0) + 50;
    cols.forEach((c, i) => {
      const d = el('div', 'hx-coll');
      d.textContent = colName(i);
      d.style.flex = '0 0 ' + c.col.width + 'px';
      d.style.width = c.col.width + 'px';
      head.appendChild(d);
      state.colHeads.push(d);
    });
    table.appendChild(head);

    // 数据行（行既可以是单元格数组，也可以是 {cells:[...]} 包装；
    // 还可以是 {group:'分类名', groupNote:'…', href:'…'} 的分组标题行）
    sheet.rows.forEach((rowCells, ri) => {
      if (!Array.isArray(rowCells) && rowCells && rowCells.group != null) {
        const gtr = el('div', 'hx-tr hx-tr-group');        const grh = el('div', 'hx-rowhead');
        grh.textContent = String(ri + 1);
        gtr.appendChild(grh);
        const gd = el('div', 'hx-cell hx-group-cell');
        gd.style.flex = '0 0 ' + (tableWidth - 50) + 'px';
        gd.style.width = (tableWidth - 50) + 'px';
        if (rowCells.href) {
          const a = el('a', 'hx-link');
          a.href = rowCells.href;
          a.textContent = rowCells.group;
          gd.appendChild(a);
        } else {
          gd.appendChild(document.createTextNode(rowCells.group));
        }
        if (rowCells.groupNote) {
          const note = el('span', 'hx-group-note', '　' + rowCells.groupNote);
          gd.appendChild(note);
        }
        gtr.appendChild(gd);
        table.appendChild(gtr);
        state.cells.push([gd]);
        return;
      }

      const arr = Array.isArray(rowCells) ? rowCells : ((rowCells && rowCells.cells) || []);
      const tr = el('div', 'hx-tr');
      const rh = el('div', 'hx-rowhead');
      rh.textContent = String(ri + 1);
      tr.appendChild(rh);

      const line = [];
      cols.forEach((c, ci) => {
        const col = c.col;
        const data = arr[c.idx] || {};
        const d = el('div', 'hx-cell' + (col.align === 'right' ? ' num' : '') +
          (col.wrap ? ' wrap' : '') + (col.content ? ' content' : ''));
        d.style.flex = '0 0 ' + col.width + 'px';
        d.style.width = col.width + 'px';

        if (data.node) {
          d.appendChild(data.node);
        } else if (data.href) {
          const a = el('a', 'hx-link');
          a.href = data.href;
          const label = data.text == null ? '' : String(data.text);
          a.textContent = label;
          d.appendChild(a);
        } else {
          d.textContent = data.text == null ? '' : String(data.text);
        }
        d.title = data.plain || (d.textContent || '').slice(0, 400);
        d._cell = { text: (data.plain || d.textContent || '').replace(/\s+/g, ' ').trim(), href: data.href || '' };
        d.dataset.r = ri;
        d.dataset.c = ci;
        line.push(d);
        tr.appendChild(d);
      });
      state.cells.push(line);
      table.appendChild(tr);
    });

    renderTabs();
    // 首行如果是分类段标题，就把光标放到第一条数据上
    let firstData = 0;
    for (let i = 0; i < state.cells.length; i++) {
      if (!(state.cells[i].length === 1 && state.cells[i][0].classList.contains('hx-group-cell'))) { firstData = i; break; }
    }
    setSel(firstData, 0, false);
    showTrail();
  }

  /** 内容计数（不把分组标题行算进去） */
  function dataRowCount(sheet) {
    return (sheet.rows || []).filter(r => Array.isArray(r) || (r && r.group == null)).length;
  }

  function renderTabs() {
    const model = state.model;
    const strip = R.tabstrip;
    strip.innerHTML = '';
    state.tabs = [];
    model.sheets.forEach((s, i) => {
      const t = el('div', 'hx-sheet-tab' + (i === state.sheet ? ' active' : ''));
      const span = el('span', '', s.name || ('Sheet' + (i + 1)));
      t.appendChild(span);
      t.title = s.name || '';
      t.addEventListener('click', () => {
        if (state.sheet === i) return;
        state.sheet = i;
        state.sel = { r: 0, c: 0 };
        renderSheet();
      });
      strip.appendChild(t);
      state.tabs.push(t);
    });

    // 分页（放在工作表标签右侧，紧挨状态栏，像 Excel 的滚动条区域）
    const pager = R.pager;
    pager.innerHTML = '';
    const pg = model.pager;
    if (pg && pg.total > 1) {
      const prev = el('a', pg.current <= 1 ? 'off' : '', '◀ 上一页');
      if (pg.current > 1) prev.href = pg.href(pg.current - 1);
      const next = el('a', pg.current >= pg.total ? 'off' : '', '下一页 ▶');
      if (pg.current < pg.total) next.href = pg.href(pg.current + 1);
      const info = el('span');
      info.innerHTML = '第 <b>' + pg.current + '</b> / ' + pg.total + ' 页';
      const input = document.createElement('input');
      input.type = 'text';
      input.value = String(pg.current);
      input.title = '跳到指定页（回车）';
      input.addEventListener('keydown', e => {
        if (e.key !== 'Enter') return;
        const n = clamp(num(input.value) || 1, 1, pg.total);
        go(pg.href(n));
      });
      const goBtn = el('button', '', 'GO');
      goBtn.addEventListener('click', () => {
        const n = clamp(num(input.value) || 1, 1, pg.total);
        go(pg.href(n));
      });
      pager.appendChild(prev);
      pager.appendChild(info);
      pager.appendChild(next);
      pager.appendChild(input);
      pager.appendChild(goBtn);
    } else {
      const st = el('span', '', '就绪');
      const count = el('span', '', '计数: ' + dataRowCount(model.sheets[state.sheet] || {}));
      pager.appendChild(st);
      pager.appendChild(count);
    }
  }

  function setSel(r, c, scroll) {
    const sheet = state.model.sheets[state.sheet];
    if (!sheet || !state.cells.length) return;
    r = clamp(r, 0, state.cells.length - 1);
    c = clamp(c, 0, state.cells[0].length - 1);
    if (state.cells[r]) c = clamp(c, 0, state.cells[r].length - 1);
    state.sel = { r: r, c: c };

    if (setSel._cell) setSel._cell.classList.remove('sel');
    if (setSel._head) setSel._head.classList.remove('sel');
    if (setSel._row) setSel._row.classList.remove('sel');

    const cell = state.cells[r] && state.cells[r][c];
    if (!cell) return;
    cell.classList.add('sel');
    const head = state.colHeads[c];
    if (head) head.classList.add('sel');
    const rowHead = cell.parentNode.querySelector('.hx-rowhead');
    if (rowHead) rowHead.classList.add('sel');

    setSel._cell = cell;
    setSel._head = head;
    setSel._row = rowHead;

    R.namebox.textContent = colName(c) + (r + 1);
    // 编辑栏不再被单元格内容顶掉：它固定显示面包屑（社区 » 版面 » … » 标题），
    // 点回复也能随时看到自己在哪、点着回上一层。单元格内容本来就在格子里。
    if (scroll !== false && cell.scrollIntoView) {
      const box = R.table.parentNode;
      const cb = cell.getBoundingClientRect();
      const bb = box.getBoundingClientRect();
      if (cb.bottom > bb.bottom - 4) box.scrollTop += cb.bottom - bb.bottom + 6;
      if (cb.top < bb.top + 26) box.scrollTop -= bb.top + 26 - cb.top;
      if (cb.right > bb.right - 4) box.scrollLeft += cb.right - bb.right + 6;
      if (cb.left < bb.left + 56) box.scrollLeft -= bb.left + 56 - cb.left;
    }
  }

  /* ============================== 8. 交互 ============================== */

  function bindRoot() {
    const root = R.root;

    // 点击单元格 / 行号 → 选中（Excel 的选中框）
    root.addEventListener('mousedown', e => {
      const cell = e.target.closest && e.target.closest('.hx-cell');
      if (cell && cell.dataset.r != null) {
        setSel(+cell.dataset.r, +cell.dataset.c, false);
        return;
      }
      const rh = e.target.closest && e.target.closest('.hx-rowhead');
      if (rh && state.cells.length) {
        const tr = rh.parentNode;
        const idx = Array.prototype.indexOf.call(tr.parentNode.children, tr) - 1;
        if (idx >= 0) setSel(idx, 0, false);
      }
    });

    // Excel 里的站内链接走软导航，不让浏览器换文档 —— 切版面 / 进帖子 / 翻页
    // 都不会再闪一下原生页面（原理见第 10 节）。外部链接（my.hupu.com、其它站）
    // 保持原样：该新标签就新标签，该整页跳转就整页跳转。
    root.addEventListener('click', e => {
      if (!CFG.enabled || isPeek()) return;
      if (e.button !== 0 || e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
      const a = e.target && e.target.closest && e.target.closest('a[href]');
      if (!a || !root.contains(a)) return;
      if (a.target && a.target !== '_self') return;
      if (!softable(a.href)) return;
      e.preventDefault();
      e.stopPropagation();
      go(a.href);
    }, true);

    // 屏蔽事件冒泡到虎扑自己的全局监听器（原生 DOM 只是藏起来，监听器都还在）
    ['click', 'mousedown', 'mouseup', 'dblclick', 'contextmenu', 'keydown', 'keyup', 'wheel', 'touchstart']
      .forEach(type => root.addEventListener(type, e => e.stopPropagation()));

    // 标题栏搜索框：回车走虎扑站内搜索
    R.search.addEventListener('keydown', e => {
      e.stopPropagation();
      if (e.key !== 'Enter') return;
      const q = R.search.value.trim();
      if (!q) return;
      go('https://bbs.hupu.com/search?q=' + encodeURIComponent(q));
    });

    // 设置 / 回首页（标题栏图标）
    const gear = one('.hx-gear', root);
    if (gear) gear.addEventListener('click', e => { e.stopPropagation(); openSettings(); });
    const home = one('.hx-home', root);
    if (home) home.addEventListener('click', e => {
      e.stopPropagation();
      if (location.pathname !== '/') go('https://bbs.hupu.com/');
    });

    bindZoom();
  }

  /* ---- 鼠标悬停浮出大图 ---- */

  function zoomPop() {
    if (!R) return null;
    let pop = one('.hx-zoom-pop', R.root);
    if (!pop) {
      pop = el('div', 'hx-zoom-pop');
      pop.innerHTML = '<img alt="" referrerpolicy="no-referrer"><span class="hx-zoom-cap"></span>';
      R.root.appendChild(pop);
    }
    return pop;
  }

  function hideZoomPop() {
    if (!R) return;
    const pop = one('.hx-zoom-pop', R.root);
    if (pop) pop.classList.remove('on');
  }

  function bindZoom() {
    const table = R.table;

    // 贴着鼠标放，贴到边就翻到另一边
    const place = (pop, ev) => {
      const w = pop.offsetWidth;
      const h = pop.offsetHeight;
      let x = ev.clientX + 18;
      let y = ev.clientY + 18;
      if (x + w > window.innerWidth - 10) x = Math.max(10, ev.clientX - w - 18);
      if (y + h > window.innerHeight - 10) y = Math.max(10, window.innerHeight - h - 10);
      pop.style.left = x + 'px';
      pop.style.top = y + 'px';
    };

    // 表格内容会重建，所以用事件委托挂在 .hx-table 上
    table.addEventListener('mouseover', e => {
      if (!CFG.imgHoverZoom) return;
      const t = e.target;
      const src = t && t.dataset ? t.dataset.zoom : '';
      if (!src) return;
      const pop = zoomPop();
      if (!pop) return;
      const img = pop.querySelector('img');
      const cap = pop.querySelector('.hx-zoom-cap');
      if (img.dataset.src !== src) {
        img.dataset.src = src;
        cap.textContent = '载入中…';
        img.onload = () => { cap.textContent = img.naturalWidth + ' × ' + img.naturalHeight; };
        img.onerror = () => { cap.textContent = '图片加载失败'; };
        img.src = src;
      }
      pop.classList.add('on');
      place(pop, e);
    });

    table.addEventListener('mousemove', e => {
      const pop = one('.hx-zoom-pop.on', R.root);
      if (pop) place(pop, e);
    });

    table.addEventListener('mouseout', e => {
      const t = e.target;
      if (t && t.dataset && t.dataset.zoom) hideZoomPop();
    });

    // 滚动后位置就不对了，直接收起
    if (table.parentNode) table.parentNode.addEventListener('scroll', hideZoomPop, true);
  }

  function bindKeys() {
    document.addEventListener('keydown', e => {
      // 老板键：Esc 立刻切回原页面 / 再按一次回来（不写存储）
      if (e.key === 'Escape') {
        if (!CFG.enabled) return;
        // 设置面板开着时，Esc 先关面板
        if (dlg) { e.preventDefault(); e.stopPropagation(); closeSettings(); return; }
        // 正在搜索框里打字时，Esc 先清空输入框，不当老板键用
        if (R && e.target === R.search && R.search.value) {
          R.search.value = '';
          return;
        }
        e.preventDefault();
        peek();
        return;
      }
      // Alt+E 开关 Excel 模式
      if (e.altKey && !e.ctrlKey && (e.key === 'e' || e.key === 'E' || e.code === 'KeyE')) {
        e.preventDefault();
        toggle();
        return;
      }
      if (!CFG.enabled || isPeek()) return;
      const t = e.target;
      if (t && t.closest && t.closest('input, textarea, [contenteditable]')) return;

      const ctrl = e.ctrlKey || e.metaKey;
      const pg = state.model && state.model.pager;

      // Ctrl+PageUp / PageDown：切换工作表
      if (ctrl && (e.key === 'PageDown' || e.key === 'PageUp')) {
        if (state.model.sheets.length > 1) {
          e.preventDefault();
          const n = state.model.sheets.length;
          state.sheet = (state.sheet + (e.key === 'PageDown' ? 1 : n - 1)) % n;
          state.sel = { r: 0, c: 0 };
          renderSheet();
          toast('工作表：' + (state.model.sheets[state.sheet].name || ''));
        }
        return;
      }
      // Alt+←/→：翻页
      if (e.altKey && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') && pg && pg.total > 1) {
        e.preventDefault();
        const n = pg.current + (e.key === 'ArrowRight' ? 1 : -1);
        if (n >= 1 && n <= pg.total) go(pg.href(n));
        return;
      }

      const sel = state.sel;
      switch (e.key) {
        case 'ArrowDown': e.preventDefault(); setSel(sel.r + 1, sel.c); return;
        case 'ArrowUp': e.preventDefault(); setSel(sel.r - 1, sel.c); return;
        case 'ArrowRight': e.preventDefault(); setSel(sel.r, sel.c + 1); return;
        case 'ArrowLeft': e.preventDefault(); setSel(sel.r, sel.c - 1); return;
        case 'PageDown': e.preventDefault(); setSel(sel.r + 20, sel.c); return;
        case 'PageUp': e.preventDefault(); setSel(Math.max(0, sel.r - 20), sel.c); return;
        case 'Home':
          e.preventDefault();
          if (ctrl) { setSel(0, 0); } else { setSel(sel.r, 0); }
          return;
        case 'End':
          e.preventDefault();
          setSel(sel.r, state.cells[sel.r] ? state.cells[sel.r].length - 1 : 0);
          return;
        case 'Enter': {
          const cell = state.cells[sel.r] && state.cells[sel.r][sel.c];
          const href = cell && cell._cell && cell._cell.href;
          if (href) { e.preventDefault(); go(href); }
          else { e.preventDefault(); setSel(sel.r + 1, sel.c); }
          return;
        }
        default: return;
      }
    }, true);
  }

  /* ============================== 8.5 设置面板 ============================== */

  /*
   * 所有可配置项集中在这里（面板里一项一行，改完立即生效并存进 GM 存储）。
   * type: bool（开关）/ enum（下拉）/ text（输入框）
   */
  const SETTINGS = [
    {
      g: '常规', key: 'enabled', type: 'bool', title: 'Excel 模式',
      desc: '关掉就直接显示虎扑原页面，和按 Alt+E 一个效果'
    },
    {
      g: '常规', key: 'theme', type: 'enum', title: '皮肤',
      options: [['office', 'Office（绿，带功能区）'], ['tencent', '腾讯文档（蓝，极简）'], ['wps', 'WPS（深蓝）']],
      desc: '切换整体配色；腾讯文档皮肤会隐藏功能区，更像在线表格'
    },
    {
      g: '常规', key: 'book', type: 'text', title: '工作簿名称',
      desc: '显示在浏览器标签和标题栏；留空则使用页面标题'
    },
    {
      g: '视图', key: 'showAccount', type: 'bool', title: '显示账号区',
      desc: '标题栏右侧的登录/消息/分享/批注；登录后会自动显示你的昵称与头像'
    },
    {
      g: '视图', key: 'showUrl', type: 'bool', title: '显示「路径」列',
      desc: '表格最后一列显示板块/帖子的路径（/topic-daily 这种）。名字本身就是链接，默认关掉更清爽'
    },
    {
      g: '视图', key: 'freezeHeader', type: 'bool', title: '冻结列标行',
      desc: '滚动表格时，列标（A、B、C…）与行号保持可见'
    },
    {
      g: '正文图片', key: 'showImages', type: 'bool', title: '显示正文图片',
      desc: '关掉后帖子里的图片会折叠成「【图片】」链接，上班摸鱼更低调'
    },
    {
      g: '正文图片', key: 'imgMaxW', type: 'range', min: 120, max: 600, step: 10, unit: 'px',
      title: '缩略图宽度上限',
      desc: '表格里图片的最大宽度，调小一点更像表格'
    },
    {
      g: '正文图片', key: 'imgMaxH', type: 'range', min: 80, max: 400, step: 10, unit: 'px',
      title: '缩略图高度上限',
      desc: '表格里图片的最大高度'
    },
    {
      g: '正文图片', key: 'imgHoverZoom', type: 'bool', title: '鼠标悬停浮出大图',
      desc: '鼠标移到缩略图上时在旁边浮出原图；原图地址会自动去掉图床的缩放参数'
    }
  ];

  const SHORTCUTS = [
    ['Alt+E', '开关 Excel 模式（会记住）'],
    ['Esc', '老板键：立刻切回虎扑原页面，再按一次回来'],
    ['↑ ↓ ← →', '像 Excel 一样移动选中的单元格'],
    ['PageUp / PageDown', '上下翻 20 行'],
    ['Home / Ctrl+Home / End', '跳到本行首列 / A1 / 本行末列'],
    ['Enter', '单元格里有链接就打开，否则下移一格'],
    ['Ctrl+PageDown / Ctrl+PageUp', '切换工作表'],
    ['Alt+← / Alt+→', '上一页 / 下一页']
  ];

  let dlg = null;

  function settingsGroups() {
    const list = [];
    SETTINGS.forEach(s => { if (list.indexOf(s.g) < 0) list.push(s.g); });
    list.push('关于');
    return list;
  }

  function settingRowHtml(s) {
    let control = '';
    if (s.type === 'bool') {
      control = '<label class="hx-switch"><input type="checkbox" data-key="' + s.key + '"' +
        (CFG[s.key] ? ' checked' : '') + '><span></span></label>';
    } else if (s.type === 'enum') {
      control = '<select data-key="' + s.key + '">' +
        s.options.map(o => '<option value="' + esc(o[0]) + '"' +
          (String(CFG[s.key]) === String(o[0]) ? ' selected' : '') + '>' + esc(o[1]) + '</option>').join('') +
        '</select>';
    } else if (s.type === 'range') {
      const v = clamp(num(CFG[s.key]), s.min, s.max);
      const pct = Math.round((v - s.min) / (s.max - s.min) * 100);
      control = '<div class="hx-range">' +
        '<input type="range" data-key="' + s.key + '" min="' + s.min + '" max="' + s.max +
        '" step="' + (s.step || 1) + '" value="' + v + '" style="--fill:' + pct + '%">' +
        '<span class="hx-range-val">' + v + esc(s.unit || '') + '</span></div>';
    } else {
      control = '<input type="text" data-key="' + s.key + '" value="' + esc(CFG[s.key]) + '" spellcheck="false">';
    }
    return '<div class="hx-set-row">' +
      '<div class="hx-set-main"><div class="hx-set-title">' + esc(s.title) + '</div>' +
      (s.desc ? '<div class="hx-set-desc">' + esc(s.desc) + '</div>' : '') + '</div>' +
      control + '</div>';
  }

  /** 启动时序的一句话总结 */
  function bootSummary() {
    if (BOOT.firstPaint < 0) return '首帧还没记录到（页面可能一直在加载）';
    const parts = ['脚本开始 ' + BOOT.script + 'ms'];
    parts.push('挂锁 ' + (BOOT.hide < 0 ? '—' : BOOT.hide + 'ms'));
    parts.push('Excel 外壳 ' + (BOOT.shell < 0 ? '—' : BOOT.shell + 'ms'));
    parts.push('首帧 ' + BOOT.firstPaint + 'ms');
    const late = BOOT.hide >= 0 && BOOT.firstPaint > 0 && BOOT.hide > BOOT.firstPaint;
    parts.push(late
      ? '⚠ 挂锁比首帧晚了 ' + (BOOT.hide - BOOT.firstPaint) + 'ms（这段是原生页面，属于油猴注入时机，见 README）'
      : '✓ 挂锁早于首帧 ' + (BOOT.firstPaint - BOOT.hide) + 'ms（首帧就是 Excel）');
    return parts.join(' · ');
  }

  /** 记录首帧时刻；如果挂锁晚于首帧，说明脚本注入太晚，给一条提示 */
  function reportBoot() {
    try {
      const paints = (performance.getEntriesByType && performance.getEntriesByType('paint')) || [];
      paints.forEach(p => {
        if (p.name === 'first-paint' && BOOT.firstPaint < 0) BOOT.firstPaint = Math.round(p.startTime);
        if (p.name === 'first-contentful-paint') BOOT.fcp = Math.round(p.startTime);
      });
      if (BOOT.firstPaint < 0 && BOOT.fcp > 0) BOOT.firstPaint = BOOT.fcp;
      if (BOOT.firstPaint < 0 || BOOT.reported) return;
      BOOT.reported = true;
      if (BOOT.hide >= 0 && BOOT.hide > BOOT.firstPaint) {
        console.warn('[hupu-excel] 脚本注入晚了 ' + (BOOT.hide - BOOT.firstPaint) + 'ms：' +
          '页面首帧在 ' + BOOT.firstPaint + 'ms，我们到 ' + BOOT.hide + 'ms 才挂上锁，' +
          '这一小段显示的是虎扑原生页面。页面里已经没法再提前（锁是脚本的第一件事），' +
          '要彻底消掉请看 README「为什么有时候会闪一下原生页面」里的浏览器级 CSS 方案。');
      }
    } catch (e) { /* 忽略 */ }
  }

  function settingsHtml() {
    const groups = settingsGroups();
    const nav = groups.map((g, i) =>
      '<div class="hx-dlg-navitem' + (i ? '' : ' active') + '" data-g="' + esc(g) + '">' + esc(g) + '</div>').join('');
    const main = groups.map((g, i) => {
      let body;
      if (g === '关于') {
        body = SHORTCUTS.map(k =>
          '<div class="hx-key-row"><kbd>' + esc(k[0]) + '</kbd><span>' + esc(k[1]) + '</span></div>').join('');
        body += '<div class="hx-set-desc" style="margin-top:12px">虎扑 Excel · 摸鱼模式 v1.0.0 —— ' +
          '只改外观，不发送任何请求、不碰账号数据。数据来自页面自身的 window.$$data / __NEXT_DATA__。</div>';
        body += '<div class="hx-set-desc" style="margin-top:8px">启动时序：' + esc(bootSummary()) + '</div>';
      } else {
        body = SETTINGS.filter(s => s.g === g).map(settingRowHtml).join('');
      }
      return '<div class="hx-set-group' + (i ? '' : ' active') + '" data-g="' + esc(g) + '">' +
        '<div class="hx-set-group-title">' + esc(g) + '</div>' + body + '</div>';
    }).join('');

    return '<div class="hx-dlg-mask"></div>' +
      '<div class="hx-dlg" role="dialog" aria-label="Excel 选项">' +
        '<div class="hx-dlg-title"><span>Excel 选项</span><i class="hx-dlg-x" data-act="close">✕</i></div>' +
        '<div class="hx-dlg-body">' +
          '<div class="hx-dlg-nav">' + nav + '</div>' +
          '<div class="hx-dlg-main">' + main + '</div>' +
        '</div>' +
        '<div class="hx-dlg-foot">' +
          '<button class="hx-dlg-btn" data-act="reset">恢复默认值</button>' +
          '<span class="hx-dlg-tip">改动即时生效并自动保存</span>' +
          '<button class="hx-dlg-btn primary" data-act="close">关闭</button>' +
        '</div>' +
      '</div>';
  }

  function openSettings() {
    if (!CFG.enabled) {
      CFG.enabled = true;
      saveCfg();
      apply();
    }
    ensureRoot();
    if (dlg && R.root.contains(dlg)) return;
    dlg = el('div', 'hx-dlg-wrap');
    dlg.innerHTML = settingsHtml();
    R.root.appendChild(dlg);
    bindSettings();
  }

  function closeSettings() {
    if (dlg) { dlg.remove(); }
    dlg = null;
  }

  function bindSettings() {
    const box = dlg;
    if (!box) return;

    $$$('.hx-dlg-navitem', box).forEach(item => {
      item.addEventListener('click', () => {
        const g = item.dataset.g;
        $$$('.hx-dlg-navitem', box).forEach(x => x.classList.toggle('active', x === item));
        $$$('.hx-set-group', box).forEach(sec => sec.classList.toggle('active', sec.dataset.g === g));
      });
    });

    $$$('[data-act="close"]', box).forEach(b => b.addEventListener('click', closeSettings));
    const mask = one('.hx-dlg-mask', box);
    if (mask) mask.addEventListener('click', closeSettings);

    const reset = one('[data-act="reset"]', box);
    if (reset) reset.addEventListener('click', () => {
      if (!confirm('把所有设置恢复成默认值？')) return;
      Object.assign(CFG, DEFAULTS);
      saveCfg();
      closeSettings();
      apply();
      toast('已恢复默认设置');
    });

    $$$('[data-key]', box).forEach(input => {
      const key = input.dataset.key;

      // 滑块：拖动时即时生效（只改 CSS 变量，不重渲染表格），松手才存
      if (input.type === 'range') {
        const spec = SETTINGS.filter(s => s.key === key)[0] || {};
        const label = input.parentNode.querySelector('.hx-range-val');
        const refresh = () => {
          const v = num(input.value);
          if (spec.max > spec.min) input.style.setProperty('--fill', Math.round((v - spec.min) / (spec.max - spec.min) * 100) + '%');
          if (label) label.textContent = v + (spec.unit || '');
          CFG[key] = v;
          applyTweaks();
        };
        input.addEventListener('input', refresh);
        input.addEventListener('change', () => { refresh(); saveCfg(); });
        return;
      }

      input.addEventListener('change', () => {
        CFG[key] = input.type === 'checkbox' ? input.checked : input.value;
        saveCfg();
        apply();
        if (!CFG.enabled) closeSettings();
        else if (key === 'book') toast('工作簿名称：' + (CFG.book || '（页面标题）'));
      });
    });
  }

  /** 标题栏账号区：登录后显示昵称 + 头像，未登录显示「登录 / 消息」 */
  function renderAccount() {
    if (!R || !R.acct) return;
    if (!CFG.showAccount) { R.acct.style.display = 'none'; return; }
    R.acct.style.display = 'flex';

    const acc = readAccount();
    if (acc) {
      R.acct.innerHTML =
        '<span class="hx-chip">' + ICONS.bell + '消息</span>' +
        '<span class="hx-chip">' + ICONS.share + '分享</span>' +
        '<span class="hx-chip">' + ICONS.comment + '批注</span>' +
        '<a class="hx-chip hx-account" href="' + esc(acc.url || 'https://my.hupu.com') + '" target="_blank" rel="noreferrer">' +
        (acc.avatar ? '<img class="hx-avatar-img" src="' + esc(acc.avatar) + '" referrerpolicy="no-referrer">'
          : '<span class="hx-avatar">' + esc(acc.name.slice(0, 1)) + '</span>') +
        '<span class="hx-uname">' + esc(acc.name) + '</span></a>';
      R.acct.title = '已登录：' + acc.name;
    } else {
      R.acct.innerHTML =
        '<span class="hx-chip">' + ICONS.user + '登录</span>' +
        '<span class="hx-chip">' + ICONS.bell + '消息</span>' +
        '<span class="hx-chip">' + ICONS.share + '分享</span>' +
        '<span class="hx-chip">' + ICONS.comment + '批注</span>' +
        '<div class="hx-avatar">虎</div>';
      R.acct.title = isLoggedIn() ? '已登录' : '未登录';
    }
  }

  /** 站点 JS 是异步把昵称写进 .hp-topLogin-info 的，变了就跟着刷新 */
  function watchAccount() {
    const box = one('.hp-topLogin-info');
    if (!box || watchAccount._done) return;
    watchAccount._done = true;
    const obs = new MutationObserver(() => {
      if (!CFG.enabled) return;
      clearTimeout(watchAccount._t);
      watchAccount._t = setTimeout(() => {
        const name = readAccount();
        const before = renderAccount._key || '';
        const now = name ? name.name : '';
        if (now !== before) { renderAccount._key = now; renderAccount(); }
      }, 300);
    });
    obs.observe(box, { childList: true, subtree: true, characterData: true });
  }

  /* ============================== 9. 开关 / 标题 / 图标 ============================== */

  let wantTitle = '';
  let origTitle = '';
  let origIcon = null;

  function isPeek() { return document.documentElement.classList.contains('hx-peek'); }

  function setTitle() {
    if (!origTitle) origTitle = document.title;
    wantTitle = CFG.book || (state.model && state.model.title) || '';
    if (wantTitle) document.title = wantTitle;
  }

  function restoreTitle() {
    if (origTitle) document.title = origTitle;
  }

  function watchTitle() {
    const target = document.querySelector('title') || document.head;
    if (!target) return;
    const obs = new MutationObserver(() => {
      if (!CFG.enabled || isPeek()) return;
      if (wantTitle && document.title !== wantTitle) document.title = wantTitle;
    });
    obs.observe(target, { childList: true, characterData: true, subtree: true });
  }

  const EXCEL_ICON = 'data:image/svg+xml,' + encodeURIComponent(
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">' +
    '<rect x="2" y="2" width="28" height="28" rx="4" fill="#217346"/>' +
    '<path d="M20.8 10.2 11.6 21.8M11.6 10.2l9.2 11.6" stroke="#ffffff" stroke-width="2.8" stroke-linecap="round" fill="none"/>' +
    '</svg>');

  function setFavicon() {
    const links = $$$('link[rel~="icon"], link[rel="shortcut icon"]');
    if (!origIcon) origIcon = links.map(l => ({ node: l, href: l.getAttribute('href') }));
    let link = links[0];
    if (!link) {
      link = document.createElement('link');
      link.rel = 'shortcut icon';
      (document.head || document.documentElement).appendChild(link);
      origIcon.push({ node: link, href: null });
    }
    link.type = 'image/svg+xml';
    link.href = EXCEL_ICON;
  }

  function restoreFavicon() {
    if (!origIcon) return;
    origIcon.forEach(it => {
      if (!it.node) return;
      if (it.href == null) it.node.remove();
      else { it.node.type = 'image/x-icon'; it.node.href = it.href; }
    });
  }

  function peek(force) {
    const next = force == null ? !isPeek() : !!force;
    // 软导航之后物理 DOM 还是旧的：先立刻切成原生页面（老板键要的就是「马上」），
    // 同时重新加载当前地址，并在新文档里直接进原生模式 —— 等新文档上来，看到的
    // 就是地址栏里那一页真正的原生页面了
    if (next && location.href !== docUrl) {
      try { sessionStorage.setItem('hx.native', '1'); } catch (e) { /* 忽略 */ }
      document.documentElement.classList.add('hx-peek');
      document.documentElement.classList.remove('hx-on');
      restoreTitle();
      restoreFavicon();
      location.replace(location.href);
      return;
    }
    document.documentElement.classList.toggle('hx-peek', next);
    // 隐藏原生页面的规则挂在 hx-on 上，peek 时把它一并摘掉，
    // 让原生页面回到站点自己的样式（而不是被我们覆盖成 display:block）
    document.documentElement.classList.toggle('hx-on', !next && !!CFG.enabled);
    if (next) { restoreTitle(); restoreFavicon(); }
    else { setTitle(); setFavicon(); }
  }

  function toggle() {
    CFG.enabled = !CFG.enabled;
    saveCfg();
    apply();
    if (CFG.enabled) toast('已开启 Excel 模式（Alt+E 关闭，Esc 快速切回）');
  }

  function apply() {
    if (CFG.enabled) {
      injectCss();
      document.documentElement.classList.add('hx-on');
      document.documentElement.classList.remove('hx-peek');
      if (BOOT.hide < 0) BOOT.hide = Math.round(performance.now());
      if (document.body) { render(); setTitle(); setFavicon(); }
    } else {
      closeSettings();
      // 关掉 Excel 模式后挂上 hx-peek：给浏览器级 CSS 一个「现在是原生模式」的信号
      document.documentElement.classList.remove('hx-on');
      document.documentElement.classList.add('hx-peek');
      if (R && R.root) { R.root.remove(); R = null; }
      restoreTitle();
      restoreFavicon();
      // 软导航过的话 DOM 是旧的，重新加载当前地址才能看到真正的原页面
      if (location.href !== docUrl) location.replace(location.href);
    }
  }

  /* ============================== 10. 启动 ============================== */

  /*
   * ── 软导航：不换文档的页面切换 ─────────────────────────────────────────────
   *
   * 虎扑是「多文档」站点：点版面、点帖子、点下一页都是整页跳转。整页跳转会创建
   * 一份新文档，而油猴只能在 document-start 注入脚本 —— 新文档从创建到第一次
   * 绘制可能只有 ~30ms，脚本还没来得及挂锁，那一小段画出来的就是原生页面
   * （README 里量过，这个窗口在页面里没法再往前压）。
   *
   * 所以 Excel 自己 UI 里的链接不再让浏览器换文档：自己 fetch 回新页面的 HTML、
   * 用 DOMParser 解析、套同一套数据读取逻辑，然后在同一个文档里重画。文档不换，
   * 原生页面就没有机会露脸。地址栏用 pushState 保持同步，前进/后退照常工作。
   *
   * 拿不到模型（搜索页、个人中心…）或请求失败时，退回原来的整页跳转。
   */

  const docUrl = location.href;   // 物理 DOM 真正对应的地址（软导航不会改 DOM）
  let viewUrl = docUrl;           // Excel 里当前展示的地址
  let viewDoc = document;         // Excel 里当前展示的那份文档（软导航时是 fetch 回来的）
  let navSeq = 0;                 // 并发软导航：只认最后一次

  /** 这个地址能不能软导航（同源、且是脚本能解析的 bbs 路由） */
  function softable(url) {
    let u;
    try { u = new URL(url, location.href); } catch (e) { return false; }
    if (u.origin !== location.origin) return false;
    if (!/(^|\.)hupu\.com$/i.test(u.hostname)) return false;
    const p = u.pathname.replace(/\/+$/, '') || '/';
    if (p === '/') return true;
    if (/^\/\d+(?:-\d+)?\.html$/.test(p)) return true;
    if (/^\/[\w-]+(?:-\d+)?$/.test(p)) return true;
    return false;
  }

  /** 用指定文档构建模型（默认当前 document）；DOC 用完还原 */
  function buildModelFor(doc) {
    const prev = DOC;
    DOC = doc || document;
    PAGE_DATA = undefined;
    try {
      return buildModel();
    } finally {
      DOC = prev;
      PAGE_DATA = undefined;
    }
  }

  function render() {
    if (!CFG.enabled || !document.body) return;
    let model = null;
    try {
      // 用「当前展示的那份文档」，而不是物理 DOM —— 否则软导航之后改任何设置
      // 都会把表格退回上一次整页加载的内容
      model = buildModelFor(viewDoc);
    } catch (err) {
      console.warn('[hupu-excel] 解析失败：', err);
    }
    if (!model || !model.sheets || !model.sheets.length) model = fallbackModel();
    ensureRoot();
    paint(model, true);   // 同一页面的重绘：留在用户原来那张表
  }

  /**
   * 软导航：地址改成 url，同时在同一份文档里重画 Excel。
   * 任何一步出问题都退回 location.href（整页跳转，行为同以前）。
   */
  function softNav(url, opts) {
    opts = opts || {};
    if (!CFG.enabled || isPeek() || !softable(url)) { location.href = url; return; }

    let target;
    try { target = new URL(url, location.href).href; } catch (e) { location.href = url; return; }
    if (target === viewUrl) return;   // 已经展示的就是这页，别多压一条历史
    if (opts.push !== false) {
      try { history.pushState(null, '', target); } catch (e) { location.href = target; return; }
    }
    lastUrl = location.href;

    const seq = ++navSeq;
    toast('正在打开 ' + (shortPath(new URL(target).pathname) || target) + ' …');
    // Accept 用普通文档的，别让服务端以为是 AJAX 而回 JSON
    fetch(target, {
      credentials: 'same-origin',
      headers: { 'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8' }
    })
      .then(res => {
        if (!res.ok) throw new Error('HTTP ' + res.status);
        return res.text().then(html => ({ html: html, finalUrl: res.url || target }));
      })
      .then(res => {
        if (seq !== navSeq) return;                 // 期间又点了一次，这次作废
        if (res.finalUrl !== location.href) {       // 跟随过重定向
          try { history.replaceState(null, '', res.finalUrl); } catch (e) { /* 忽略 */ }
          lastUrl = location.href;
        }
        const doc = new DOMParser().parseFromString(res.html, 'text/html');
        let model = null;
        try { model = buildModelFor(doc); } catch (e) { console.warn('[hupu-excel] 软导航解析失败：', e); }
        if (!model || !model.sheets || !model.sheets.length) throw new Error('这个页面没有可摊开的数据');
        viewDoc = doc;
        viewUrl = location.href;
        ensureRoot();
        paint(model);
        setTitle();
      })
      .catch(err => {
        if (seq !== navSeq) return;
        console.warn('[hupu-excel] 软导航失败，改用整页跳转：', err);
        location.href = target;
      });
  }

  /** 脚本自己 UI 里的跳转统一走这里：能软导航就软导航，否则整页跳转 */
  function go(url) { softNav(url); }

  function hookHistory() {
    ['pushState', 'replaceState'].forEach(k => {
      const orig = history[k];
      if (typeof orig !== 'function') return;
      history[k] = function () {
        const ret = orig.apply(this, arguments);
        setTimeout(onUrlChange, 0);
        return ret;
      };
    });
    window.addEventListener('popstate', onUrlChange);
    window.addEventListener('hashchange', onUrlChange);
  }

  let lastUrl = location.href;
  function onUrlChange() {
    if (location.href === lastUrl) return;
    lastUrl = location.href;
    if (!CFG.enabled || isPeek()) return;
    // 前进/后退：当前位置不是软导航自己切的，重新拉一份渲染
    if (softable(location.href)) softNav(location.href, { push: false });
    else if (location.href === docUrl) { viewDoc = document; viewUrl = docUrl; render(); }   // 回到物理 DOM 本身就是的那页
    else location.reload();                       // 其它情况重新加载当前地址
  }

  function onReady() {
    injectCss();
    if (CFG.enabled) {
      render();
      renderAccount();
      watchAccount();
      // 这次是「加载后直接显示原生页面」（软导航后按 Esc 触发的重载）：
      // 外壳照建（Esc 再按一次就能看到 Excel），但不抢标题 / 图标
      if (!START_NATIVE) { setTitle(); setFavicon(); watchTitle(); }
      // 账号区是站点 JS 异步渲染的，晚一点再补看一次
      setTimeout(() => { if (CFG.enabled) { watchAccount(); renderAccount(); } }, 2500);
      if (!onReady._hinted && !START_NATIVE) {
        onReady._hinted = true;
        toast('Excel 模式已开启 · Alt+E 关闭 · Esc 切回原页面 · 右上角 ⚙ 设置', 5600);
      }
    }
    BOOT.domReady = Math.round(performance.now());
    // 首帧时刻要等第一帧画完才有（DOMContentLoaded 时可能还没画），
    // 所以分几次采集；面板里的「启动时序」每次打开都会读最新的值
    reportBoot();
    requestAnimationFrame(() => {
      reportBoot();
      [120, 600, 1600].forEach(ms => setTimeout(reportBoot, ms));
    });
    booted = true;
  }

  /* ------------------------------ 启动 ------------------------------ */

  /** 越早越好：先把原生页面藏起来，避免闪一下再变成 Excel */
  function earlyStart() {
    injectCss();
    if (CFG.enabled && !START_NATIVE) {
      document.documentElement.classList.add('hx-on');
      if (BOOT.hide < 0) BOOT.hide = Math.round(performance.now());
    } else {
      // 没开 Excel 模式（或这次要直接显示原生页面）时挂上 hx-peek：这样浏览器级
      // CSS（见 README）只要写 html:not(.hx-peek) 一条就能既挡首帧、又不会在
      // 关掉脚本时把页面藏没
      document.documentElement.classList.add('hx-peek');
    }
  }

  /**
   * <body> 一出现就把 Excel 外壳（标题栏/功能区/编辑栏/「正在打开 工作簿1.xlsx…」）
   * 立起来，真正的数据等 DOMContentLoaded 再填。
   *
   * 为什么不等 DOMContentLoaded 一起做：真实站点从首帧到 DOMContentLoaded 有
   * 1～2 秒（首页那次测到 1.1 秒），这段时间里如果只挂了 hx-on，用户看到的是
   * 一片白；如果那个页面的内容不在 #container / #__next 里，看到的直接就是
   * 原生页面。先立外壳，首帧就是 Excel。
   */
  function earlyShell() {
    if (!CFG.enabled) return;
    earlyStart();
    if (!document.body) return;
    if (BOOT.shell < 0) BOOT.shell = Math.round(performance.now());
    ensureRoot();
    renderAccount();
    applyTweaks();
    if (!START_NATIVE) { setTitle(); setFavicon(); }
  }

  /**
   * 兜底看门狗：万一站点把我们加的 hx-on / #hx-root 弄掉了（或脚本注入晚了），
   * 这里再把状态拉回来，保证「要么是 Excel，要么什么都没有」。
   */
  function guard() {
    if (!CFG.enabled) {
      // 关掉的时候也要替浏览器级 CSS 维持好「原生模式」这个信号
      if (!document.documentElement.classList.contains('hx-peek')) {
        document.documentElement.classList.add('hx-peek');
      }
      return;
    }
    if (!document.documentElement.classList.contains('hx-on') && !isPeek()) {
      document.documentElement.classList.add('hx-on');
    }
    if (isPeek() || !booted) return;
    if (!R || !document.body || !document.body.contains(R.root)) {
      if (document.body) {
        try { render(); } catch (e) { /* 下次再看 */ }
      }
    }
  }

  /** onReady 跑完了没（看门狗在启动过程中不要抢着渲染） */
  let booted = false;

  function installGuard() {
    const check = () => { try { guard(); } catch (e) {} };
    document.addEventListener('readystatechange', check);
    window.addEventListener('pageshow', check);
    window.addEventListener('focus', check);
    [400, 1200, 3000].forEach(ms => setTimeout(check, ms));
    // 站点动了我们的 class / 把 #hx-root 删了：下一个微任务就补回来，
    // 不用等定时器（等的话那段时间里原生页面已经画出来了）
    whenDocumentElement(() => {
      const obs = new MutationObserver(check);
      obs.observe(document.documentElement, { childList: true, attributes: true, attributeFilter: ['class'] });
      whenBody(() => obs.observe(document.body, { childList: true }));
    });
  }

  whenDocumentElement(earlyStart);
  whenBody(earlyShell);
  hookHistory();
  bindKeys();
  installGuard();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', onReady);
  } else {
    onReady();
  }
})();
