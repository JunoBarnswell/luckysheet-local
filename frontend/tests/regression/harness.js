(function (root) {
  var results = [];
  var start = Date.now();

  function el(id) {
    return document.getElementById(id);
  }

  function log(ok, name, detail) {
    results.push({ ok: ok, name: name, detail: detail || "" });
    var box = el("assert-log");
    if (!box) box = el("report");
    if (!box) return;
    var line = document.createElement("div");
    line.className = ok ? "pass" : "fail";
    line.textContent = (ok ? "PASS" : "FAIL") + "  " + name + (detail ? " — " + detail : "");
    box.appendChild(line);
  }

  function assert(cond, name, detail) {
    if (cond) log(true, name, detail);
    else log(false, name, detail);
    return !!cond;
  }

  function assertEqual(actual, expected, name) {
    var ok = actual === expected;
    log(ok, name, "actual=" + JSON.stringify(actual) + " expected=" + JSON.stringify(expected));
    return ok;
  }

  function summary() {
    var pass = results.filter(function (r) { return r.ok; }).length;
    var fail = results.length - pass;
    var head = el("assert-summary");
    if (!head) head = el("report");
    if (head) {
      head.textContent = "断言 " + pass + " 通过 / " + fail + " 失败 / 共 " + results.length +
        "（" + (Date.now() - start) + "ms）";
      head.className = fail ? "fail" : "pass";
    }
    return { pass: pass, fail: fail, results: results };
  }

  function resolveLuckysheetSrc() {
    var candidates = [
      "../../dist/luckysheet.umd.js",
      "../../src/luckysheet.umd.js",
      "../../../frontend/dist/luckysheet.umd.js"
    ];
    var q = /[?&]lib=([^&]+)/.exec(location.search);
    if (q) candidates.unshift(decodeURIComponent(q[1]));
    return candidates;
  }

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = src;
      s.onload = function () { resolve(src); };
      s.onerror = function () { reject(new Error(src)); };
      document.head.appendChild(s);
    });
  }

  function loadCss(href) {
    var l = document.createElement("link");
    l.rel = "stylesheet";
    l.href = href;
    document.head.appendChild(l);
  }

  function tryLoadLuckysheet() {
    loadCss("../../dist/plugins/css/pluginsCss.css");
    loadCss("../../dist/plugins/plugins.css");
    loadCss("../../dist/css/luckysheet.css");
    var plugin = loadScript("../../dist/plugins/js/plugin.js").catch(function () { return null; });
    return plugin.then(function () {
      var list = resolveLuckysheetSrc();
      var i = 0;
      function next() {
        if (i >= list.length) {
          return Promise.reject(new Error("未找到 luckysheet.umd.js，请先在 frontend 执行 npm run build"));
        }
        var src = list[i++];
        return loadScript(src).then(function () {
          if (!root.luckysheet) throw new Error("loaded but no window.luckysheet");
          return src;
        }).catch(next);
      }
      return next();
    });
  }

  function wait(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms); });
  }

  function cellV(r, c) {
    var cell = root.luckysheet.getCellValue(r, c, { type: "v" });
    return cell;
  }

  root.LsRegression = {
    assert: assert,
    assertEqual: assertEqual,
    log: log,
    summary: summary,
    tryLoadLuckysheet: tryLoadLuckysheet,
    wait: wait,
    cellV: cellV
  };
})(window);
