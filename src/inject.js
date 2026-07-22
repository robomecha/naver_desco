/*
 * inject.js  — 페이지 컨텍스트(world:MAIN)에서 실행됩니다.
 * 네이버 부동산이 호출하는 fetch/XHR 응답을 가로채 매물 JSON을 content script 로 전달하고,
 * 상세 API 재호출에 필요한 인증 헤더(Authorization 등)를 기억합니다.
 * 또한 Next.js 서버렌더링 데이터(window.__NEXT_DATA__, self.__next_f)도 함께 넘겨줍니다.
 */
(function () {
  "use strict";

  if (window.__NRE_INJECTED__) return;
  window.__NRE_INJECTED__ = true;
  try { console.log("[매물정리:inject] MAIN world 로드됨"); } catch (e) {}

  var SOURCE = "NRE_PAGE";
  // 정적 자원은 응답 본문을 읽지 않음
  var ASSET_RE = /\.(js|mjs|css|png|jpe?g|gif|svg|webp|avif|woff2?|ttf|otf|ico|map|mp4|webm|mp3)(\?|#|$)/i;
  // 인증 헤더를 기억할 대상(네이버 부동산 API 요청)
  var AUTH_RE = /(land\.naver\.com|front-api|\/api\/|cluster)/i;

  var capturedHeaders = {};
  var stats = { resp: 0, article: 0 };

  function rememberHeaders(headers) {
    if (!headers) return;
    try {
      Object.keys(headers).forEach(function (k) {
        var lk = k.toLowerCase();
        if (lk === "authorization" || lk.indexOf("x-") === 0) capturedHeaders[k] = headers[k];
      });
    } catch (e) {}
  }

  function headersToObject(h) {
    var out = {};
    if (!h) return out;
    try {
      if (typeof Headers !== "undefined" && h instanceof Headers) h.forEach(function (v, k) { out[k] = v; });
      else if (Array.isArray(h)) h.forEach(function (p) { out[p[0]] = p[1]; });
      else if (typeof h === "object") Object.keys(h).forEach(function (k) { out[k] = h[k]; });
    } catch (e) {}
    return out;
  }

  function post(type, payload) {
    try { window.postMessage({ __nre: true, source: SOURCE, type: type, payload: payload }, "*"); } catch (e) {}
  }

  function handleResponseText(url, text) {
    if (!text) return;
    var t = text.length > 3000000 ? text.slice(0, 3000000) : text;
    var s = t.replace(/^﻿/, "");
    var c = s.charAt(0);
    if (c !== "{" && c !== "[") return; // JSON 만
    var data;
    try { data = JSON.parse(s); } catch (e) { return; }
    stats.resp++;
    post("apiResponse", { url: url, data: data });
  }

  function shouldReadBody(url) {
    if (!url) return true;         // 상대경로 등 URL 불명확 → 일단 읽음
    if (ASSET_RE.test(url)) return false;
    return true;
  }

  /* ---------- fetch 패치 ---------- */
  var origFetch = window.fetch;
  if (origFetch) {
    window.fetch = function (input, init) {
      var url = "";
      try { url = typeof input === "string" ? input : (input && input.url) || ""; } catch (e) {}
      if (AUTH_RE.test(url)) {
        try { rememberHeaders(headersToObject((init && init.headers) || (input && input.headers))); } catch (e) {}
      }
      var p = origFetch.apply(this, arguments);
      if (shouldReadBody(url)) {
        p.then(function (res) {
          try {
            var ct = "";
            try { ct = res.headers.get("content-type") || ""; } catch (e) {}
            if (!url && ct && ct.indexOf("json") < 0 && ct.indexOf("text") < 0) return res;
            res.clone().text().then(function (tx) { handleResponseText(url, tx); }).catch(function () {});
          } catch (e) {}
          return res;
        }).catch(function () {});
      }
      return p;
    };
  }

  /* ---------- XMLHttpRequest 패치 ---------- */
  var XHR = window.XMLHttpRequest;
  if (XHR) {
    var origOpen = XHR.prototype.open;
    var origSend = XHR.prototype.send;
    var origSetHeader = XHR.prototype.setRequestHeader;

    XHR.prototype.open = function (method, url) {
      this.__nre_url = url;
      this.__nre_read = shouldReadBody(url);
      this.__nre_auth = AUTH_RE.test(url || "");
      this.__nre_headers = {};
      return origOpen.apply(this, arguments);
    };
    XHR.prototype.setRequestHeader = function (name, value) {
      try { if (this.__nre_auth) this.__nre_headers[name] = value; } catch (e) {}
      return origSetHeader.apply(this, arguments);
    };
    XHR.prototype.send = function () {
      if (this.__nre_auth) rememberHeaders(this.__nre_headers);
      if (this.__nre_read) {
        var self = this;
        this.addEventListener("load", function () {
          try {
            var t = null;
            if (self.responseType === "" || self.responseType === "text") t = self.responseText;
            else if (self.responseType === "json") t = JSON.stringify(self.response);
            if (t) handleResponseText(self.__nre_url, t);
          } catch (e) {}
        });
      }
      return origSend.apply(this, arguments);
    };
  }

  /* ---------- Next.js 서버렌더링 데이터 스캔 ---------- */
  var nextDataPosted = false;
  var rscPostedUpto = 0; // 이미 보낸 RSC 청크 개수 (같은 청크 반복 전송 방지)

  function scanNextData() {
    try {
      if (!nextDataPosted && window.__NEXT_DATA__) {
        nextDataPosted = true;
        post("apiResponse", { url: "__NEXT_DATA__", data: window.__NEXT_DATA__ });
      }
    } catch (e) {}
    // RSC 스트리밍 청크(self.__next_f) 중 '새로 추가된 것만' 전달
    try {
      var f = window.__next_f;
      if (f && f.length > rscPostedUpto) {
        for (var i = rscPostedUpto; i < f.length; i++) {
          var chunk = f[i];
          var str = Array.isArray(chunk) ? chunk[chunk.length - 1] : chunk;
          if (typeof str === "string" && /atclNo|articleNo|articleNumber|articleList|realEstateType/.test(str)) {
            post("rscChunk", { text: str });
          }
        }
        rscPostedUpto = f.length;
      }
    } catch (e) {}
  }

  window.addEventListener("message", function (ev) {
    if (ev.source !== window) return;
    var msg = ev.data;
    if (!msg || !msg.__nre || msg.source !== "NRE_CONTENT") return;

    if (msg.type === "rescan") {
      scanNextData();
      post("stats", stats);
    } else if (msg.type === "enrichFetch") {
      var req = msg.payload || {};
      var reqId = req.reqId;
      var headers = { accept: "application/json" };
      Object.keys(capturedHeaders).forEach(function (k) { headers[k] = capturedHeaders[k]; });
      (origFetch || window.fetch).call(window, req.url, { method: "GET", headers: headers, credentials: "include" })
        .then(function (res) {
          return res.text().then(function (t) {
            var data = null; try { data = JSON.parse(t); } catch (e) {}
            post("enrichResult", { reqId: reqId, ok: res.ok, status: res.status, url: req.url, data: data });
          });
        })
        .catch(function (err) { post("enrichResult", { reqId: reqId, ok: false, status: 0, url: req.url, error: String(err) }); });
    }
  });

  // 로드 시점별로 서버렌더링 데이터 스캔
  if (document.readyState === "complete" || document.readyState === "interactive") scanNextData();
  document.addEventListener("DOMContentLoaded", scanNextData);
  window.addEventListener("load", function () { setTimeout(scanNextData, 800); });

  post("ready", { hasFetch: !!origFetch, hasXHR: !!XHR });
})();
