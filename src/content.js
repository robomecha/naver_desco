/*
 * content.js — 네이버부동산 매물정리 표 (콘텐츠 스크립트)
 * 1) inject.js 를 페이지에 삽입해 API 응답을 가로챕니다.
 * 2) 매물 데이터를 정규화/누적하고, 화면에 떠 있는 표 패널로 보여줍니다.
 * 3) 주소·사용승인년도가 없는 매물은 상세 API 로 자동 보강합니다.
 * 4) CSV / 엑셀(xls) / 클립보드로 내보냅니다.
 */
(function () {
  "use strict";

  var HOST = location.hostname; // fin.land.naver.com / new.land.naver.com / www.disco.re
  var IS_FIN = /fin\.land\.naver\.com/.test(HOST);
  var IS_DISCO = /disco\.re/.test(HOST);
  var PYEONG = 3.305785; // 1평 = 3.305785 ㎡

  /* ================= 데이터 저장소 ================= */
  var articles = new Map();   // articleNo -> record
  var dismissed = new Set();  // 초기화로 지운 매물번호 (재스캔으로 되살아나지 않게)
  var renderTimer = null;
  var lastCount = 0;
  // 사용자가 버튼을 눌러 패널을 처음 열기 전에는 아무것도 수집하지 않음
  var collectionEnabled = false;

  /* inject.js 는 manifest 의 world:"MAIN" 콘텐츠 스크립트로 페이지에 직접 주입됩니다. */
  try { console.log("[매물정리] content 로드됨 @", location.href); } catch (e) {}

  /* ================= 유틸: 값 추출 ================= */
  function firstOf(obj, keys) {
    for (var i = 0; i < keys.length; i++) {
      var v = obj[keys[i]];
      if (v !== undefined && v !== null && v !== "") return v;
    }
    return undefined;
  }

  // 객체 트리를 순회하며 조건에 맞는 첫 값 반환
  function deepFind(root, predicate, maxDepth) {
    maxDepth = maxDepth == null ? 6 : maxDepth;
    var stack = [{ node: root, depth: 0 }];
    while (stack.length) {
      var cur = stack.pop();
      var node = cur.node;
      if (!node || typeof node !== "object" || cur.depth > maxDepth) continue;
      var keys = Object.keys(node);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i];
        var val = node[k];
        if (val !== null && typeof val !== "object") {
          try {
            if (predicate(k, val)) return val;
          } catch (e) {}
        }
      }
      for (var j = 0; j < keys.length; j++) {
        var cv = node[keys[j]];
        if (cv && typeof cv === "object") stack.push({ node: cv, depth: cur.depth + 1 });
      }
    }
    return undefined;
  }

  /* ================= 매물 객체 판별 & 정규화 ================= */
  var ARTICLE_NO_KEYS = ["articleNumber", "articleNo", "atclNo", "articleId", "atclNoStr"];
  function getArticleNo(o) {
    var v = firstOf(o, ARTICLE_NO_KEYS);
    return v == null ? null : String(v);
  }
  function looksLikeArticle(o) {
    if (!o || typeof o !== "object") return false;
    if (getArticleNo(o) == null) return false;
    // 거래유형/가격/면적 정보 중 하나라도 있으면 매물로 간주
    return (
      firstOf(o, ["tradeTypeName", "tradTpNm", "tradeTypeCode", "tradeType", "realEstateType"]) != null ||
      firstOf(o, ["priceInfo", "dealOrWarrantPrc", "prc", "dealPrice", "warrantPrice", "price"]) != null ||
      firstOf(o, ["spaceInfo", "spaceInfoDto", "area1", "spc1", "supplyArea"]) != null
    );
  }

  /* 코드 → 한글 표기 */
  var TRADE_MAP = { A1: "매매", B1: "전세", B2: "월세", B3: "단기임대" };
  var RET_MAP = {
    A01: "아파트", A02: "오피스텔", A03: "재건축", A04: "아파트분양권", A05: "오피스텔분양권",
    B01: "빌라/연립", B02: "연립", B03: "다세대",
    C01: "단독/다가구", C02: "상가주택", C03: "한옥주택", C04: "전원주택",
    D01: "원룸", D02: "상가", D03: "사무실", D04: "공장/창고", D05: "건물", D06: "재개발",
    E03: "토지", E01: "원룸", E02: "고시원", G01: "지식산업센터"
  };
  function tradeName(code, fallback) { return TRADE_MAP[code] || fallback || code || ""; }
  function retName(code, fallback) { return RET_MAP[code] || fallback || code || ""; }

  function manwonFromWon(v) {
    var n = toNum(v);
    if (n == null || n === 0) return null;
    return Math.round(n / 10000);
  }

  /* 주소 객체를 사람이 읽는 문자열로 조립 */
  function assembleAddress(addr) {
    if (addr == null) return "";
    if (typeof addr === "string") return addr;
    if (typeof addr !== "object") return "";
    // 완성형 주소 필드 우선
    var pref = firstOf(addr, ["roadAddress", "jibunAddress", "fullAddress", "fullRoadAddress", "roadAddressName"]);
    if (pref && typeof pref === "string") return pref;
    var parts = [];
    Object.keys(addr).forEach(function (k) {
      if (/code|id$|Id|type|Type|url|Url|lat|lng|Lat|Lng|(^|_)x$|(^|_)y$|number$|Number$/.test(k)) return;
      var v = addr[k];
      if (typeof v === "string" && v && !/^\d+$/.test(v)) {
        if (parts.indexOf(v) < 0) parts.push(v);
      }
    });
    return parts.join(" ").trim();
  }

  // 주소를 문자열로: 완전한 주소 문자열(지번 포함)이 있으면 그걸, 없으면 컴포넌트 조립(동까지).
  function addressToString(addr) {
    if (addr == null) return "";
    if (typeof addr === "string") return addr;
    if (typeof addr !== "object") return "";
    var full = bestAddressString(addr);
    if (full && /\d/.test(full)) return full;      // 번지 있는 완전한 주소 우선
    var asm = assembleAddress(addr);
    return (full && full.length > asm.length) ? full : asm;
  }

  // 객체 안에서 '가장 완전한' 한국 주소 문자열을 찾음(지번/도로명 포함). 키 이름에 의존하지 않음.
  function bestAddressString(obj, maxDepth) {
    maxDepth = maxDepth == null ? 5 : maxDepth;
    var best = "";
    var stack = [{ node: obj, depth: 0 }], seen = 0;
    while (stack.length && seen < 4000) {
      var cur = stack.pop(), n = cur.node; seen++;
      if (!n || typeof n !== "object" || cur.depth > maxDepth) continue;
      var keys = Object.keys(n);
      for (var i = 0; i < keys.length; i++) {
        var k = keys[i], v = n[k];
        if (typeof v === "string") {
          if (/code|url|id$|Id/i.test(k)) continue;
          // 한국 주소로 보이는 문자열: 행정구역/도로 단어 포함 + 여러 토큰(공백)이나 번지(숫자)
          var isAddr = /[가-힣]/.test(v) && v.length >= 5 &&
            /(시|도|군|구|읍|면|동|리|가|로|길)/.test(v) && (/\s/.test(v) || /\d/.test(v));
          if (isAddr) {
            var score = v.length + (/\d/.test(v) ? 100 : 0); // 숫자(번지) 있으면 가점
            var bestScore = best.length + (/\d/.test(best) ? 100 : 0);
            if (score > bestScore) best = v;
          }
        } else if (v && typeof v === "object" && cur.depth < maxDepth) {
          stack.push({ node: v, depth: cur.depth + 1 });
        }
      }
    }
    return best;
  }

  // 주소가 동(洞)까지만 있고 번지가 없는지 (지번 보강 필요 여부)
  function needsJibun(rec) {
    if (!rec.address) return true;
    return !/\d/.test(String(rec.address)); // 숫자(번지)가 없으면 동까지만인 것으로 간주
  }

  function deepFindObject(root, keyPred, maxDepth) {
    maxDepth = maxDepth == null ? 6 : maxDepth;
    var stack = [{ node: root, depth: 0 }];
    while (stack.length) {
      var cur = stack.pop(), node = cur.node;
      if (!node || typeof node !== "object" || cur.depth > maxDepth) continue;
      var keys = Object.keys(node);
      for (var i = 0; i < keys.length; i++) {
        var v = node[keys[i]];
        if (v && typeof v === "object" && !Array.isArray(v)) {
          try { if (keyPred(keys[i], v)) return v; } catch (e) {}
        }
      }
      for (var j = 0; j < keys.length; j++) {
        var cv = node[keys[j]];
        if (cv && typeof cv === "object") stack.push({ node: cv, depth: cur.depth + 1 });
      }
    }
    return null;
  }

  // JSON 트리 전체에서 매물 객체들을 수집
  function collectArticles(root, out) {
    var stack = [{ node: root, depth: 0 }];
    var seen = 0;
    while (stack.length && seen < 20000) {
      var cur = stack.pop();
      var node = cur.node;
      seen++;
      if (!node || typeof node !== "object" || cur.depth > 8) continue;
      if (Array.isArray(node)) {
        for (var i = 0; i < node.length; i++) {
          if (node[i] && typeof node[i] === "object") stack.push({ node: node[i], depth: cur.depth + 1 });
        }
      } else {
        if (looksLikeArticle(node)) out.push(node);
        var keys = Object.keys(node);
        for (var j = 0; j < keys.length; j++) {
          var cv = node[keys[j]];
          if (cv && typeof cv === "object") stack.push({ node: cv, depth: cur.depth + 1 });
        }
      }
    }
  }

  function toNum(v) {
    if (v == null) return null;
    if (typeof v === "number") return v;
    var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
    return isNaN(n) ? null : n;
  }

  // "3억 5,000" / "9억" / "5,000" -> 만원 단위 숫자
  function parseManwon(str) {
    if (str == null) return null;
    if (typeof str === "number") return str; // 이미 숫자면 만원으로 간주하지 않고 그대로
    var s = String(str).replace(/\s/g, "");
    if (!s || s === "-") return null;
    var total = 0, matched = false;
    var eok = s.match(/(\d+(?:,\d+)?)\s*억/);
    if (eok) {
      total += parseFloat(eok[1].replace(/,/g, "")) * 10000;
      matched = true;
      s = s.replace(/.*억/, "");
    }
    var man = s.match(/(\d[\d,]*)/);
    if (man) {
      var m = parseFloat(man[1].replace(/,/g, ""));
      if (!isNaN(m)) { total += m; matched = true; }
    }
    return matched ? Math.round(total) : null;
  }

  // 만원 -> "3억 5,000" 표기
  function formatManwon(manwon) {
    if (manwon == null || isNaN(manwon)) return "";
    manwon = Math.round(manwon);
    if (manwon === 0) return "0";
    var eok = Math.floor(manwon / 10000);
    var rest = manwon % 10000;
    var parts = [];
    if (eok > 0) parts.push(eok + "억");
    if (rest > 0) parts.push(rest.toLocaleString("ko-KR"));
    return parts.join(" ").trim();
  }

  function extractYear(v) {
    if (v == null) return null;
    var s = String(v);
    var m = s.match(/(19|20)\d{2}/);
    return m ? m[0] : null;
  }

  /* ================= 진단: 매물로 의심되는 객체 구조 수집 ================= */
  var debugSamples = [];       // {sig, keys, sample}
  var debugSigs = {};
  var DEBUG_ID_RE = /atcl|article|realestate|complex/i;
  var DEBUG_VAL_RE = /price|prc|deal|rent|warrant|trade|area|spc|floor|flr|addr|주소|승인|approve|aprv/i;

  function trimVal(v) {
    if (v == null) return v;
    if (typeof v === "string") return v.length > 40 ? v.slice(0, 40) + "…" : v;
    if (typeof v === "object") return Array.isArray(v) ? "[array:" + v.length + "]" : "{object}";
    return v;
  }

  function debugScan(root) {
    if (debugSamples.length >= 40) return;
    var stack = [{ node: root, depth: 0 }], seen = 0;
    while (stack.length && seen < 20000) {
      var cur = stack.pop(); seen++;
      var node = cur.node;
      if (!node || typeof node !== "object" || cur.depth > 8) continue;
      if (Array.isArray(node)) {
        for (var a = 0; a < node.length; a++) if (node[a] && typeof node[a] === "object") stack.push({ node: node[a], depth: cur.depth + 1 });
        continue;
      }
      var keys = Object.keys(node);
      if (keys.length > 45) { // 지도 레이어 등 거대한 객체는 노이즈 → 기록 생략, 하위 탐색만
        for (var kk = 0; kk < keys.length; kk++) {
          var cvv = node[keys[kk]];
          if (cvv && typeof cvv === "object") stack.push({ node: cvv, depth: cur.depth + 1 });
        }
        continue;
      }
      var hasId = false, hasVal = false;
      for (var i = 0; i < keys.length; i++) {
        if (DEBUG_ID_RE.test(keys[i])) hasId = true;
        if (DEBUG_VAL_RE.test(keys[i])) hasVal = true;
      }
      if (hasId || hasVal) {
        var sig = keys.slice(0, 30).join(",");
        if (!debugSigs[sig] && debugSamples.length < 40) {
          debugSigs[sig] = true;
          var sample = {};
          for (var j = 0; j < keys.length && j < 40; j++) {
            var val = node[keys[j]];
            // 주소·좌표 객체는 한 단계 펼쳐서 내부 필드를 보여줌(지번 확인용)
            if (val && typeof val === "object" && !Array.isArray(val) && /address|주소|coordinate|소재/i.test(keys[j])) {
              var inner = {};
              Object.keys(val).slice(0, 30).forEach(function (ik) { inner[ik] = trimVal(val[ik]); });
              sample[keys[j]] = inner;
            } else {
              sample[keys[j]] = trimVal(val);
            }
          }
          debugSamples.push({ keys: keys, sample: sample });
        }
      }
      for (var k = 0; k < keys.length; k++) {
        var cv = node[keys[k]];
        if (cv && typeof cv === "object") stack.push({ node: cv, depth: cur.depth + 1 });
      }
    }
  }

  // 리스트/상세 객체를 표준 레코드로 변환(부분 병합용)
  function normalize(o) {
    var no = getArticleNo(o);
    if (!no) return null;
    var rec = { articleNo: no };

    // 종류/거래유형 (코드 + 한글)
    var retCode = firstOf(o, ["realEstateType", "realEstateTypeCode", "rletTpCd"]);
    var ttCode = firstOf(o, ["tradeType", "tradeTypeCode", "tradTpCd"]);
    if (retCode) rec.realEstateType = retCode;
    if (ttCode) rec.tradeType = ttCode;
    var retNm = firstOf(o, ["realEstateTypeName", "rletTpNm"]);
    var ttNm = firstOf(o, ["tradeTypeName", "tradTpNm"]);
    rec.realEstateTypeName = retName(retCode, retNm);
    rec.tradeTypeName = tradeName(ttCode, ttNm);
    rec.name = firstOf(o, ["articleName", "atclNm", "buildingName", "complexName", "bildNm"]);

    // 가격 (fin.land: priceInfo 객체, 단위 '원')
    var pi = o.priceInfo && typeof o.priceInfo === "object" ? o.priceInfo : (o.priceInfoDto && typeof o.priceInfoDto === "object" ? o.priceInfoDto : o);
    var deal = manwonFromWon(firstOf(pi, ["dealPrice"]));
    var warrant = manwonFromWon(firstOf(pi, ["warrantyPrice", "warrantPrice"]));
    var rentP = manwonFromWon(firstOf(pi, ["rentPrice"]));
    if (deal != null) rec.dealManwon = deal;
    if (warrant != null) rec.depositManwon = warrant;
    if (rentP != null) rec.rentManwon = rentP;
    // 구(舊) new.land 문자열 형식 대비
    if (rec.dealManwon == null && rec.depositManwon == null) {
      var legacy = parseManwon(firstOf(o, ["dealOrWarrantPrc", "prc", "price"]));
      if (legacy != null) {
        if (/매매/.test(rec.tradeTypeName)) rec.dealManwon = legacy;
        else rec.depositManwon = legacy;
      }
      if (rec.rentManwon == null) rec.rentManwon = parseManwon(firstOf(o, ["rentPrc"]));
    }

    // 면적 (fin.land: spaceInfo/spaceInfoDto, 단위 ㎡)
    var sp = o.spaceInfo && typeof o.spaceInfo === "object" ? o.spaceInfo : (o.spaceInfoDto && typeof o.spaceInfoDto === "object" ? o.spaceInfoDto : o);
    var a1 = toNum(firstOf(sp, ["supplySpace", "area1", "spc1", "supplyArea"]));
    var a2 = toNum(firstOf(sp, ["exclusiveSpace", "area2", "spc2", "exclusiveArea"]));
    var al = toNum(firstOf(sp, ["landSpace", "lotSpace"]));
    if (a1 != null && a1 > 0) rec.area1 = a1;     // 공급 ㎡
    if (a2 != null && a2 > 0) rec.area2 = a2;     // 전용 ㎡
    if (al != null && al > 0) rec.landArea = al;  // 토지 ㎡ (토지 매물용)

    // 층/방향/확인일 (articleDetail / verificationInfo 안에 있을 수 있음)
    var ad = o.articleDetail && typeof o.articleDetail === "object" ? o.articleDetail : o;
    var floor = firstOf(ad, ["floorInfo", "flrInfo", "floor"]);
    if (floor) rec.floor = floor;
    var dir = firstOf(ad, ["direction", "directionName", "dirNm"]);
    if (dir) rec.direction = dir;
    var vi = o.verificationInfo && typeof o.verificationInfo === "object" ? o.verificationInfo : o;
    var cfm = firstOf(vi, ["articleConfirmDate", "articleConfirmYmd", "atclCfmYmd", "cfmYmd", "exposureStartDate"]);
    if (cfm) rec.confirmYmd = cfm;

    // 주소 (fin.land: address 객체) — 지번 포함 완전한 주소 우선
    var addr = firstOf(o, ["address", "roadAddress", "exposureAddress", "cortarAddress"]);
    var addrStr = addressToString(addr);
    if (addrStr) rec.address = addrStr;

    // 좌표 (fin.land: coordinates.xCoordinate=경도, yCoordinate=위도, WGS84)
    var co = deepFindObject(o, function (k, v) {
      return /coordinate/i.test(k) && (v.xCoordinate != null || v.yCoordinate != null || v.lat != null);
    });
    if (co) {
      var lng = toNum(firstOf(co, ["xCoordinate", "longitude", "lng", "lon", "x"]));
      var lat = toNum(firstOf(co, ["yCoordinate", "latitude", "lat", "y"]));
      if (lat != null && Math.abs(lat) < 90) rec.lat = lat;
      if (lng != null && Math.abs(lng) < 180) rec.lng = lng;
    }

    // 사용승인일 (article 안 buildingInfo 등에 있을 수 있음)
    var useY = deepFind(o, function (k, v) {
      return /useApprovalDate|useApprove|usageApprov|useAprv|approveYmd|approvalDate|사용승인|준공/i.test(k) && extractYear(v) != null;
    });
    if (useY) rec.useApproveYmd = useY;
    // 사용승인 경과년차 (네이버가 직접 제공: approvalElapsedYear)
    var el = deepFind(o, function (k, v) {
      return /approvalElapsedYear|elapsedYear|경과년/i.test(k) && v != null && !isNaN(v);
    });
    if (el != null) rec.elapsedYear = toNum(el);

    // 단지번호(있으면 단지 정보로 사용승인/주소 보강)
    var cno = firstOf(o, ["complexNumber", "complexNo", "hscpNo"]);
    if (cno != null) rec.complexNumber = String(cno);

    rec.featureDesc = firstOf(ad, ["articleFeatureDescription", "articleFeatureDesc", "atclFetrDesc"]);
    return rec;
  }

  /* 단지(complex) 정보 저장: 사용승인일·주소 보강용 */
  var complexes = new Map(); // complexNumber -> {address, useApproveYmd, name}
  function absorbComplexes(root) {
    var stack = [{ node: root, depth: 0 }], seen = 0;
    while (stack.length && seen < 20000) {
      var cur = stack.pop(), node = cur.node; seen++;
      if (!node || typeof node !== "object" || cur.depth > 8) continue;
      if (Array.isArray(node)) {
        for (var a = 0; a < node.length; a++) if (node[a] && typeof node[a] === "object") stack.push({ node: node[a], depth: cur.depth + 1 });
        continue;
      }
      var cno = firstOf(node, ["complexNumber", "complexNo"]);
      if (cno != null && (node.useApprovalDate || node.approvalElapsedYear != null || node.address || node.name)) {
        var key = String(cno);
        var c = complexes.get(key) || {};
        if (node.useApprovalDate) c.useApproveYmd = node.useApprovalDate;
        if (node.approvalElapsedYear != null && !isNaN(node.approvalElapsedYear)) c.elapsedYear = toNum(node.approvalElapsedYear);
        if (node.name) c.name = node.name;
        var ca = addressToString(node.address);
        if (ca) c.address = ca;
        complexes.set(key, c);
        // 이미 수집된 매물에 소급 반영
        applyComplexToArticles(key, c);
      }
      var keys = Object.keys(node);
      for (var k = 0; k < keys.length; k++) {
        var cv = node[keys[k]];
        if (cv && typeof cv === "object") stack.push({ node: cv, depth: cur.depth + 1 });
      }
    }
  }
  function applyComplexToArticles(cno, c) {
    articles.forEach(function (rec) {
      if (rec.complexNumber !== cno) return;
      if (!rec.useApproveYmd && c.useApproveYmd) rec.useApproveYmd = c.useApproveYmd;
      if (rec.elapsedYear == null && c.elapsedYear != null) rec.elapsedYear = c.elapsedYear;
      applyComplexAddress(rec, c);
    });
  }
  // 단지 주소로 매물 주소를 채우거나, 단지 주소에 지번이 있으면 동까지만인 주소를 업그레이드
  function applyComplexAddress(rec, c) {
    if (!c.address) return;
    var complexHasJibun = /\d/.test(c.address);
    if (!rec.address) rec.address = c.address;
    else if (complexHasJibun && needsJibun(rec)) rec.address = c.address;
  }

  function mergeRecord(rec) {
    if (!rec || !rec.articleNo) return false;
    if (dismissed.has(rec.articleNo)) return false; // 초기화로 지운 매물은 다시 담지 않음
    var existing = articles.get(rec.articleNo);
    if (!existing) {
      articles.set(rec.articleNo, rec);
      return true;
    }
    var changed = false;
    Object.keys(rec).forEach(function (k) {
      var v = rec[k];
      if (v === undefined || v === null || v === "") return;
      if (existing[k] === undefined || existing[k] === null || existing[k] === "") {
        existing[k] = v;
        changed = true;
      }
    });
    return changed;
  }

  /* ================= 페이지 메시지 수신 ================= */
  window.addEventListener("message", function (ev) {
    if (ev.source !== window) return;
    var msg = ev.data;
    if (!msg || !msg.__nre || msg.source !== "NRE_PAGE") return;

    if (msg.type === "apiResponse") {
      stats.resp++;
      handleApiData(msg.payload && msg.payload.data, msg.payload && msg.payload.url);
    } else if (msg.type === "rscChunk") {
      handleRscChunk(msg.payload && msg.payload.text);
    } else if (msg.type === "enrichResult") {
      handleEnrichResult(msg.payload);
    } else if (msg.type === "stats") {
      // 페이지측 통계(참고용)
    } else if (msg.type === "ready") {
      requestRescan();
    }
  });

  var stats = { resp: 0 };

  function requestRescan() {
    // 1) 페이지 HTML 직접 스캔(확실한 경로)
    try { scanDom(); } catch (e) {}
    // 2) MAIN world inject 에도 재스캔 요청(fetch/RSC 런타임 데이터)
    try {
      window.postMessage({ __nre: true, source: "NRE_CONTENT", type: "rescan" }, "*");
    } catch (e) {}
  }

  /* ---------- RSC(서버렌더링) 청크에서 매물 JSON 추출 ---------- */
  function extractBalanced(s, start) {
    var open = s.charAt(start), close = open === "{" ? "}" : "]";
    var depth = 0, inStr = false, esc = false;
    for (var i = start; i < s.length; i++) {
      var ch = s.charAt(i);
      if (inStr) {
        if (esc) esc = false;
        else if (ch === "\\") esc = true;
        else if (ch === '"') inStr = false;
        continue;
      }
      if (ch === '"') { inStr = true; continue; }
      if (ch === "{" || ch === "[") depth++;
      else if (ch === "}" || ch === "]") { depth--; if (depth === 0) return s.slice(start, i + 1); }
    }
    return null;
  }

  function handleRscChunk(text) {
    if (!text || typeof text !== "string") return;
    var i = 0, guard = 0;
    while (i < text.length && guard < 5000) {
      guard++;
      var b = -1;
      for (var j = i; j < text.length; j++) {
        var c = text.charAt(j);
        if (c === "{" || c === "[") { b = j; break; }
      }
      if (b < 0) break;
      var frag = extractBalanced(text, b);
      if (!frag) { i = b + 1; continue; }
      i = b + frag.length;
      if (frag.indexOf("atclNo") < 0 && frag.indexOf("articleNo") < 0 && frag.indexOf("articleId") < 0 && frag.indexOf("realEstateType") < 0) continue;
      var data;
      try { data = JSON.parse(frag); } catch (e) { continue; }
      handleApiData(data);
    }
  }

  /* ---------- 페이지 HTML 스크립트 직접 스캔 (MAIN world 미의존, 가장 확실) ---------- */
  var domScanSig = "";
  function scanDom() {
    var scripts;
    try { scripts = document.querySelectorAll("script"); } catch (e) { return; }
    // 스크립트 내용이 그대로면 재파싱 생략(성능).
    // body.innerHTML 로 재면 우리 패널이 그려질 때마다 값이 바뀌어 매번 재파싱됨 → script 기준으로 계산.
    var total = 0;
    for (var s = 0; s < scripts.length; s++) total += (scripts[s].textContent || "").length;
    var sig = scripts.length + ":" + total;
    if (sig === domScanSig) return;
    domScanSig = sig;
    var hitText = [];
    for (var i = 0; i < scripts.length; i++) {
      var t = scripts[i].textContent;
      if (!t || t.length < 8) continue;
      var id = scripts[i].id || "";

      // 1) __NEXT_DATA__ (Pages Router)
      if (id === "__NEXT_DATA__") {
        try { stats.resp++; handleApiData(JSON.parse(t)); } catch (e) {}
        continue;
      }

      // 2) App Router RSC 스트리밍: self.__next_f.push([1,"...json..."])
      if (t.indexOf("__next_f.push") >= 0) {
        var idx = 0, guard = 0;
        while (guard < 2000) {
          guard++;
          var p = t.indexOf("__next_f.push(", idx);
          if (p < 0) break;
          var br = t.indexOf("[", p);
          if (br < 0) break;
          var arrStr = extractBalanced(t, br);
          if (!arrStr) { idx = br + 1; continue; }
          idx = br + arrStr.length;
          try {
            var arr = JSON.parse(arrStr);
            var s = Array.isArray(arr) ? arr[arr.length - 1] : arr;
            if (typeof s === "string") { stats.resp++; handleRscChunk(s); }
          } catch (e) {}
        }
        continue;
      }

      // 3) 그 외 매물번호가 들어간 인라인 스크립트
      if (t.indexOf("atclNo") >= 0 || t.indexOf("articleNo") >= 0 || t.indexOf("articleId") >= 0) {
        hitText.push(t);
      }
    }
    for (var j = 0; j < hitText.length; j++) { stats.resp++; handleRscChunk(hitText[j]); }
  }

  function handleApiData(data, url) {
    if (!data) return;
    // 패널을 열기 전에는 수집하지 않음. 단, 디스코는 실거래(short_info/real_price_*) API 자체가
    // 사용자가 필지·실거래내역을 실제로 클릭했을 때만 호출되므로(자동 로드 요청은 이미 URL로 걸러냄),
    // 패널을 늦게 열더라도 그 사이의 클릭을 놓치지 않도록 이 게이트를 적용하지 않는다.
    if (!collectionEnabled && !IS_DISCO) return;
    if (IS_DISCO) { handleDiscoData(data, url); return; }
    var found = [];
    collectArticles(data, found);
    if (found.length && !window.__nreLoggedSample) {
      window.__nreLoggedSample = true;
      try {
        console.log("[매물정리] 매물 객체 예시 필드:", Object.keys(found[0]), found[0]);
      } catch (e) {}
    }
    if (!found.length) { try { debugScan(data); } catch (e) {} }
    else { try { probeArticleAddress(found); } catch (e) {} }
    // 단지(complex) 정보(사용승인일·주소) 수집
    try { absorbComplexes(data); } catch (e) {}
    var any = false;
    for (var i = 0; i < found.length; i++) {
      var rec = normalize(found[i]);
      // 단지 정보가 이미 있으면 즉시 보강
      if (rec && rec.complexNumber && complexes.has(rec.complexNumber)) {
        var c = complexes.get(rec.complexNumber);
        if (!rec.useApproveYmd && c.useApproveYmd) rec.useApproveYmd = c.useApproveYmd;
        if (rec.elapsedYear == null && c.elapsedYear != null) rec.elapsedYear = c.elapsedYear;
        applyComplexAddress(rec, c);
      }
      if (mergeRecord(rec)) any = true;
    }
    // 주의: 여기서 응답 전체를 뒤져 주소를 가져오면 안 됩니다.
    // 한 응답에 무관한 단지/매물이 섞여 오기 때문에, 엉뚱한 주소가 붙습니다.
    // 주소·사용승인은 normalize()(매물 객체 내부) 와 complexNumber 매칭으로만 채웁니다.
    if (any || articles.size !== lastCount) scheduleRender();
    if (any) { try { enqueueGeocode(); } catch (e) {} } // 새 매물 지번 자동 변환
  }

  /* ================= 디스코(disco.re) 실거래 수집 ================= */
  // 실거래 레코드 판별: 거래일 + 거래구분(gbn) + 금액(price/deposit/monthly) 중 하나
  function looksLikeDeal(o) {
    if (!o || typeof o !== "object" || Array.isArray(o)) return false;
    var hasDate = o.full_date != null || o.date != null || o.deal_date != null || o.contract_date != null;
    var hasAmt = o.price != null || o.deposit != null || o.monthly != null || o.deal_price != null;
    if (!hasDate || !hasAmt) return false;
    return o.gbn != null || o.floor != null || o.address != null || o.name != null;
  }

  function collectDeals(root, out) {
    var stack = [{ node: root, depth: 0 }], seen = 0;
    while (stack.length && seen < 20000) {
      var cur = stack.pop(), node = cur.node; seen++;
      if (!node || typeof node !== "object" || cur.depth > 8) continue;
      if (Array.isArray(node)) {
        for (var a = 0; a < node.length; a++) if (node[a] && typeof node[a] === "object") stack.push({ node: node[a], depth: cur.depth + 1 });
        continue;
      }
      if (looksLikeDeal(node)) out.push(node);
      var keys = Object.keys(node);
      for (var k = 0; k < keys.length; k++) {
        var cv = node[keys[k]];
        if (cv && typeof cv === "object") stack.push({ node: cv, depth: cur.depth + 1 });
      }
    }
  }

  // gbn/금액으로 거래유형 판별. price>0 매매, monthly>0 월세, deposit>0 전세.
  function discoTradeName(o) {
    var price = toNum(o.price != null ? o.price : o.deal_price);
    var deposit = toNum(o.deposit);
    var monthly = toNum(o.monthly);
    if (price && price > 0) return "매매";
    if (monthly && monthly > 0) return "월세";
    if (deposit && deposit > 0) return "전세";
    if (o.gbn === 1) return "매매";
    if (o.gbn === 2) return "전세";
    if (o.gbn === 3) return "월세";
    return "";
  }

  // 디스코 type 코드 = estate_type (디스코 JS SaleModel에서 확정 추출)
  var DISCO_TYPE_MAP = { 0: "단독/다가구", 1: "토지", 2: "연립/다세대", 3: "아파트", 4: "오피스텔", 5: "상업용건물", 7: "상가/사무실", 8: "공장/창고", 12: "숙박시설" };

  function normalizeDisco(o) {
    var addr = o.address || o.addr || "";
    var name = o.name || o.apt_name || o.building_name || "";
    var floor = o.floor != null ? String(o.floor) : "";
    var date = o.full_date || o.deal_date || o.contract_date || o.date || "";
    var pnu = o.pnu || o.uuid || "";
    // 같은 거래를 유일하게 식별
    var no = [pnu, date, floor, o.price, o.deposit, o.monthly, o.ho || ""].join("|");
    var rec = { articleNo: no };
    rec.address = addr;
    rec.name = name;
    rec.tradeTypeName = discoTradeName(o);
    rec.realEstateTypeName = DISCO_TYPE_MAP[o.type] || (o.type != null ? "종류" + o.type : "");
    var price = toNum(o.price != null ? o.price : o.deal_price);
    var deposit = toNum(o.deposit);
    var monthly = toNum(o.monthly);
    if (price && price > 0) rec.dealManwon = price;      // disco 금액은 만원 단위
    if (deposit && deposit > 0) rec.depositManwon = deposit;
    if (monthly && monthly > 0) rec.rentManwon = monthly;
    rec.floor = floor;
    rec.confirmYmd = date;  // 거래일(정렬/표시)
    // 면적: 전용면적(건물) / 토지면적(토지)
    var a2 = toNum(firstOf(o, ["exclusive_area", "area", "jeon_area", "jeonyong", "exclu_area", "spc2"]));
    var a1 = toNum(firstOf(o, ["supply_area", "gong_area", "spc1"]));
    var la = toNum(o.land_area);
    if (a2 && a2 > 0) rec.area2 = a2;
    if (a1 && a1 > 0) rec.area1 = a1;
    if (la && la > 0) rec.landArea = la;
    // 디스코가 직접 준 평단가(만원/평): 대지 평단가 우선, 없으면 연면적 평단가
    var spp = toNum(o.site_py_price), tpp = toNum(o.tfloor_py_price);
    if (spp && spp > 0) rec.discoPyeong = spp;
    else if (tpp && tpp > 0) rec.discoPyeong = tpp;
    var lat = toNum(o.lat), lng = toNum(o.lng);
    if (lat) rec.lat = lat; if (lng) rec.lng = lng;
    rec.dealCount = o.cnt;
    return rec;
  }

  // 지역 최근실거래(여러 필지 자동 로드) 엔드포인트는 제외 — 클릭한 실거래내역만 수집
  var DISCO_AUTO_RE = /get_recent_real_price|get_area_|get_near_|display_ad|namecard|marker/i;
  var discoDump = []; // 디스코 진단: {url, distinctPnu, count, sample}

  function distinctPnuCount(list) {
    var s = {};
    for (var i = 0; i < list.length; i++) { var p = list[i].pnu || list[i].uuid || ""; s[p] = 1; }
    return Object.keys(s).length;
  }

  // 클릭한 필지/건물의 short_info(용도지역·지목·실거래가 등)를 표준 레코드로
  function normalizeDiscoShort(o, pnu) {
    var rec = { articleNo: "disco|" + (pnu || o.address || "") };
    rec.address = o.address || o.road_address || "";
    rec.name = o.name || "";
    var isLand = (o.is_land === 1 || o.is_land === "1");
    rec.realEstateTypeName = isLand ? "토지" : (DISCO_TYPE_MAP[o.type] || (o.type != null ? "종류" + o.type : ""));
    rec.useDistrict = o.use_district || "";   // 용도지역
    rec.jimok = o.purpose || "";              // 지목
    var price = toNum(o.price);
    if (price && price > 0) { rec.dealManwon = price; rec.tradeTypeName = "매매"; } // 대표 실거래가(만원)
    rec.confirmYmd = o.realp_date || "";       // 최근 실거래일
    // 면적: 토지(land_area) / 연면적(floor_area) / 전용(exclusive_area) 별도 보관
    var la = toNum(o.land_area), ea = toNum(o.exclusive_area), fa = toNum(o.floor_area);
    rec.ziphap = (o.ziphap === 1 || o.ziphap === "1"); // 집합건물(아파트·오피스텔 등)
    if (la && la > 0) rec.landArea = la;       // 토지면적
    if (fa && fa > 0) rec.floorArea = fa;      // 연면적(건물)
    if (ea && ea > 0) rec.area2 = ea;          // 전용(집합건물 세대)
    // 평단가 기준: 집합건물=전용, 그 외(토지·일반건물)=대지(토지)
    var basis = (ea && ea > 0) ? ea : (la && la > 0 ? la : null);
    if (price && price > 0 && basis) rec.discoPyeong = price / (basis / PYEONG);
    if (o.const_year) rec.useApproveYmd = o.const_year;  // 사용승인/건축년도
    // 층수(건물 총 층수)
    if (o.ground_floor) rec.floor = "지상" + o.ground_floor + (o.basement_floor ? "/지하" + o.basement_floor : "") + "층";
    rec.pnu = pnu;
    return rec;
  }

  // 직전에 클릭한 필지/건물(short_info)의 정보 — 개별 실거래내역 클릭 시 주소·용도지역 등을 채우는 데 씀
  var lastDiscoContext = null;

  // 실거래내역 한 건(real_price_detail 또는 real_price_table 의 행)을 표준 레코드로.
  // 이 응답들엔 주소·용도지역이 없어서, 직전 클릭한 필지(lastDiscoContext)의 정보를 함께 사용.
  function buildDiscoDealRecord(row, ctx, no) {
    ctx = ctx || {};
    var rec = { articleNo: "disco-deal|" + no };
    rec.address = ctx.address || "";
    rec.name = ctx.name || "";
    rec.realEstateTypeName = ctx.realEstateTypeName || DISCO_TYPE_MAP[row.type] || "";
    rec.useDistrict = ctx.useDistrict || "";
    rec.jimok = ctx.jimok || "";
    rec.ziphap = ctx.ziphap;
    rec.tradeTypeName = discoTradeName(row);
    var price = toNum(row.price), deposit = toNum(row.deposit), monthly = toNum(row.monthly);
    if (price && price > 0) rec.dealManwon = price;
    if (deposit && deposit > 0) rec.depositManwon = deposit;
    if (monthly && monthly > 0) rec.rentManwon = monthly;
    rec.confirmYmd = row.d || row.full_date || "";
    var ea = toNum(row.exclusive_area);
    if (ea && ea > 0) rec.area2 = ea;                 // 거래마다 전용면적이 다를 수 있음(호실별)
    if (ctx.landArea) rec.landArea = ctx.landArea;    // 건물(대지)은 직전 필지 컨텍스트에서
    if (ctx.floorArea) rec.floorArea = ctx.floorArea;
    if (ctx.useApproveYmd) rec.useApproveYmd = ctx.useApproveYmd;
    rec.floor = row.floor != null ? String(row.floor) : "";
    var basis = (ea && ea > 0) ? ea : (ctx.landArea || null);
    if (rec.dealManwon != null && basis) rec.discoPyeong = rec.dealManwon / (basis / PYEONG);
    if (ctx.lat != null) rec.lat = ctx.lat;
    if (ctx.lng != null) rec.lng = ctx.lng;
    return rec;
  }

  function handleDiscoData(data, url) {
    url = url || "";
    if (!data || typeof data !== "object") return;

    // 진단용: 디스코 응답을 URL과 함께 기록
    if (discoDump.length < 10) {
      var samp = {};
      Object.keys(data).slice(0, 40).forEach(function (k) { samp[k] = trimVal(data[k]); });
      discoDump.push({ url: url.replace(/[?&]_=\d+/, "").slice(0, 180), keys: Object.keys(data), sample: samp });
    }

    // 1) 클릭한 필지/건물 = short_info (용도지역·지목·실거래가 등). 이후 실거래내역 클릭의 컨텍스트로도 씀.
    if (/\/short_info\//.test(url) && data.address) {
      var pnu = (url.match(/[?&]p=([^&]+)/) || [])[1] || "";
      var rec = normalizeDiscoShort(data, pnu);
      lastDiscoContext = rec;
      if (!window.__nreLoggedSample) { window.__nreLoggedSample = true; try { console.log("[실거래] short_info:", data); } catch (e) {} }
      if (rec && mergeRecord(rec)) {
        scheduleRender();
        try { enqueueAddrGeocode(); } catch (e) {} // 지도 내보내기용 좌표 채우기
      }
      return;
    }

    // 2) "실거래" 탭에서 개별 실거래내역 항목을 클릭하면 호출됨 (요청하신 "클릭할 때마다 수집"의 핵심)
    if (/\/real_price_detail\//.test(url)) {
      var no = (url.match(/[?&]no=([^&]+)/) || [])[1];
      if (no && data && typeof data === "object" && Object.keys(data).length) {
        var dealRec = buildDiscoDealRecord(data, lastDiscoContext, no);
        if (mergeRecord(dealRec)) {
          scheduleRender();
          try { enqueueAddrGeocode(); } catch (e) {}
        }
      }
      return;
    }

    // 3) 실거래내역 목록(여러 건이 한 번에 오는 경우)도 같은 방식으로 수집
    if (/\/real_price_table\//.test(url) && data && Array.isArray(data.real_price)) {
      var any3 = false;
      for (var i = 0; i < data.real_price.length; i++) {
        var row = data.real_price[i];
        if (!row || row.no == null) continue;
        if (mergeRecord(buildDiscoDealRecord(row, lastDiscoContext, row.no))) any3 = true;
      }
      if (any3) {
        scheduleRender();
        try { enqueueAddrGeocode(); } catch (e) {}
      }
      return;
    }
    // 그 외(get_similar 유사거래·get_recent 지역최근·마커·광고)는 수집하지 않음.
  }

  // 매물 1건의 상세(basicInfo) 응답에서만 사용. 응답 전체가 해당 매물 것이라는 전제.
  function extractDetailPatch(data, no) {
    var patch = { articleNo: no };
    var got = false;

    // fin.land: address 는 객체 → 조립. 구형: 문자열.
    var addrStr = "";
    var addrObj = deepFindObject(data, function (k) { return /^address$|addressInfo/i.test(k); });
    if (addrObj) addrStr = assembleAddress(addrObj);
    if (!addrStr) {
      var addr = deepFind(data, function (k, v) {
        if (typeof v !== "string" || v.length < 4) return false;
        return /roadAddress|exposureAddress|detailAddress|(^|_)address$|cortarAddress/i.test(k);
      });
      if (addr) addrStr = addr;
    }
    if (addrStr) { patch.address = addrStr; got = true; }

    var useY = deepFind(data, function (k, v) {
      return /useApprovalDate|useApprove|usageApprov|useAprv|approveYmd|approvalDate|사용승인|준공|cmpltYmd/i.test(k) &&
        extractYear(v) != null;
    });
    if (useY) { patch.useApproveYmd = useY; got = true; }

    return got ? patch : null;
  }

  /* ================= 상세정보 보강(자동 호출) ================= */
  var enrichWaiters = {};
  var enrichSeq = 0;

  function pageFetch(url) {
    return new Promise(function (resolve) {
      var reqId = "r" + (++enrichSeq) + "_" + Date.now();
      enrichWaiters[reqId] = resolve;
      window.postMessage({ __nre: true, source: "NRE_CONTENT", type: "enrichFetch", payload: { reqId: reqId, url: url } }, "*");
      setTimeout(function () {
        if (enrichWaiters[reqId]) {
          delete enrichWaiters[reqId];
          resolve({ ok: false, timeout: true });
        }
      }, 12000);
    });
  }

  function handleEnrichResult(payload) {
    if (!payload) return;
    var cb = enrichWaiters[payload.reqId];
    if (cb) {
      delete enrichWaiters[payload.reqId];
      cb(payload);
    }
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }

  /* ================= 카카오 좌표→지번 역지오코딩 ================= */
  // 기본 카카오 REST API 키(개인용). 저장된 키가 있으면 그걸 우선 사용.
  var DEFAULT_KAKAO_KEY = "fe4c726d3c93f04b354ed5be458c9c94";
  var kakaoKey = DEFAULT_KAKAO_KEY;
  try {
    chrome.storage && chrome.storage.local.get("kakaoKey", function (r) {
      kakaoKey = (r && r.kakaoKey) || DEFAULT_KAKAO_KEY;
    });
  } catch (e) {}

  /* ---------- 자동 지번 변환 (좌표 캐시 + 큐) ---------- */
  var autoGeocode = true;      // 자동 변환 on/off
  var geoDisabled = false;     // 401/403 등으로 중단됨(수동 🗺 또는 토글로 재개)
  var geoCache = new Map();    // "lat,lng" -> 지번 문자열(실패 시 "")
  var geoQueue = [];
  var geoActive = false;
  var geoStats = { done: 0, filled: 0 };

  function coordKey(rec) {
    return (+rec.lat).toFixed(6) + "," + (+rec.lng).toFixed(6);
  }

  // 같은 좌표의 모든 매물에 지번 적용
  function applyGeoToSameCoord(key, jibun) {
    articles.forEach(function (r) {
      if (r.lat == null || r.lng == null) return;
      if (coordKey(r) !== key) return;
      if (jibun && needsJibun(r)) r.address = jibun;
      r._geoDone = true;
    });
  }

  function enqueueGeocode() {
    if (!autoGeocode || geoDisabled || !kakaoKey) return;
    articles.forEach(function (rec) {
      if (rec.lat == null || rec.lng == null || rec._geoDone || !needsJibun(rec)) return;
      var key = coordKey(rec);
      if (geoCache.has(key)) { applyGeoToSameCoord(key, geoCache.get(key)); return; }
      if (geoQueue.indexOf(rec) < 0) geoQueue.push(rec);
    });
    runGeoQueue();
  }

  async function runGeoQueue() {
    if (geoActive) return;
    geoActive = true;
    while (geoQueue.length && !geoDisabled) {
      var rec = geoQueue.shift();
      if (!rec || rec._geoDone || rec.lat == null || !needsJibun(rec)) continue;
      var key = coordKey(rec);
      if (geoCache.has(key)) { applyGeoToSameCoord(key, geoCache.get(key)); scheduleRender(); continue; }
      var resp = await reverseGeocodeRow(rec);
      if (resp && resp.ok) {
        var jibun = resp.jibun || resp.road || "";
        geoCache.set(key, jibun);
        applyGeoToSameCoord(key, jibun);
        geoStats.done++; if (jibun) geoStats.filled++;
        scheduleRender();
      } else if (resp && (resp.status === 401 || resp.status === 403)) {
        geoDisabled = true; // API 미활성화/키 문제 → 자동 중단
        scheduleRender();
        break;
      } else {
        rec._geoDone = true; // 이 매물은 포기
      }
      await sleep(140); // 호출 간격
    }
    geoActive = false;
  }

  /* ---------- 디스코: 주소→좌표 (지도 내보내기용) ---------- */
  var addrGeoCache = new Map(); // 주소 -> {lat,lng}
  var addrGeoQueue = [];
  var addrGeoActive = false;
  function enqueueAddrGeocode() {
    if (geoDisabled || !kakaoKey) return;
    articles.forEach(function (rec) {
      if (!rec.address || (rec.lat != null && rec.lng != null) || rec._addrGeoDone) return;
      if (addrGeoCache.has(rec.address)) {
        var c = addrGeoCache.get(rec.address);
        if (c) { rec.lat = c.lat; rec.lng = c.lng; }
        rec._addrGeoDone = true; return;
      }
      if (addrGeoQueue.indexOf(rec) < 0) addrGeoQueue.push(rec);
    });
    runAddrGeoQueue();
  }
  async function runAddrGeoQueue() {
    if (addrGeoActive) return;
    addrGeoActive = true;
    while (addrGeoQueue.length && !geoDisabled) {
      var rec = addrGeoQueue.shift();
      if (!rec || rec._addrGeoDone || !rec.address || (rec.lat != null && rec.lng != null)) continue;
      var resp = await new Promise(function (resolve) {
        try { chrome.runtime.sendMessage({ type: "forwardGeocode", address: rec.address, key: kakaoKey }, function (r) { resolve(r || { ok: false }); }); }
        catch (e) { resolve({ ok: false }); }
      });
      if (resp && resp.ok) {
        addrGeoCache.set(rec.address, resp.lat != null ? { lat: resp.lat, lng: resp.lng } : null);
        if (resp.lat != null) { rec.lat = resp.lat; rec.lng = resp.lng; }
        rec._addrGeoDone = true;
        scheduleRender();
      } else if (resp && (resp.status === 401 || resp.status === 403)) {
        geoDisabled = true; break;
      } else { rec._addrGeoDone = true; }
      await sleep(140);
    }
    addrGeoActive = false;
  }

  var kakaoDump = []; // 진단용: 최근 카카오 호출 결과
  function reverseGeocodeRow(rec) {
    return new Promise(function (resolve) {
      try {
        chrome.runtime.sendMessage(
          { type: "reverseGeocode", lat: rec.lat, lng: rec.lng, key: kakaoKey },
          function (resp) {
            var le = null;
            try { le = chrome.runtime.lastError ? chrome.runtime.lastError.message : null; } catch (e) {}
            if (le && !resp) resolve({ ok: false, error: "runtime:" + le });
            else resolve(resp || { ok: false, error: "no-response" });
          }
        );
      } catch (e) { resolve({ ok: false, error: String(e) }); }
    });
  }

  var geocoding = false;
  async function fillJibunKakao(btn) {
    if (geocoding) return;
    // 키 확인/입력
    if (!kakaoKey) {
      var k = prompt(
        "카카오 REST API 키를 입력하세요.\n\n발급: developers.kakao.com → 내 애플리케이션 → 앱 → 앱 키 → 'REST API 키'\n(무료, 좌표→지번 변환에만 사용, 이 브라우저에만 저장됩니다)",
        ""
      );
      if (!k) return;
      kakaoKey = k.trim();
      try { chrome.storage.local.set({ kakaoKey: kakaoKey }); } catch (e) {}
    }
    geoDisabled = false; // 수동 실행은 중단 상태를 해제하고 재시도
    var targets = [];
    articles.forEach(function (rec) {
      if (rec.lat != null && rec.lng != null && needsJibun(rec)) targets.push(rec);
    });
    if (!targets.length) {
      flash(btn, articles.size ? "채울 좌표 없음" : "매물 없음");
      return;
    }
    geocoding = true;
    var done = 0, ok = 0, keyBad = false, apiDisabled = false;
    for (var i = 0; i < targets.length; i++) {
      var rec = targets[i];
      if (rec._geoDone || !needsJibun(rec)) { done++; continue; }
      var key = coordKey(rec);
      if (geoCache.has(key)) { // 같은 좌표는 캐시 재사용(호출 안 함)
        applyGeoToSameCoord(key, geoCache.get(key));
        if (geoCache.get(key)) ok++;
        done++; continue;
      }
      var resp = await reverseGeocodeRow(rec);
      if (kakaoDump.length < 5) kakaoDump.push({ lat: rec.lat, lng: rec.lng, resp: resp });
      if (resp && resp.ok) {
        var jibun = resp.jibun || resp.road || "";
        geoCache.set(key, jibun);
        applyGeoToSameCoord(key, jibun);
        if (jibun) ok++;
      } else if (resp && resp.status === 401) {
        keyBad = true; break;         // 키 자체가 잘못됨
      } else if (resp && resp.status === 403) {
        apiDisabled = true; break;    // 키는 맞지만 카카오맵(Local) API 미활성화
      }
      done++;
      if (btn) btn.textContent = "지번 변환 중… " + done + "/" + targets.length;
      scheduleRender();
      await sleep(120); // 카카오 호출 간격
    }
    geocoding = false;
    if (apiDisabled || keyBad) geoDisabled = true; // 자동 변환도 중단
    if (btn) btn.textContent = "🗺 지번 지금";
    if (keyBad) {
      kakaoKey = "";
      try { chrome.storage.local.remove("kakaoKey"); } catch (e) {}
      alert("카카오 API 키가 올바르지 않습니다(401).\nREST API 키가 맞는지 확인해 주세요.");
    } else if (apiDisabled) {
      // 키는 유지. 카카오맵 API 활성화 안내.
      alert(
        "카카오 지도 API가 아직 활성화되지 않았습니다(403).\n\n" +
        "developers.kakao.com → 내 애플리케이션 → (해당 앱)\n" +
        "→ 왼쪽 '카카오맵' → '카카오맵 사용 설정'을 ON 으로 바꾸고\n" +
        "저장한 뒤 다시 시도해 주세요. (무료)"
      );
    } else {
      flash(btn, ok + "건 지번 채움");
    }
  }

  var enriching = false;
  async function enrichAll(btn) {
    if (enriching) return;
    enriching = true;
    var targets = [];
    articles.forEach(function (rec) {
      // 주소 없음 / 사용승인 없음 / 동까지만(지번 없음) 이면 상세 조회
      if (!rec.address || !rec.useApproveYmd || needsJibun(rec)) targets.push(rec);
    });
    if (!targets.length) {
      enriching = false;
      flash(btn, "보강할 매물 없음");
      return;
    }
    var done = 0;
    for (var i = 0; i < targets.length; i++) {
      var rec = targets[i];
      try {
        await enrichOne(rec);
      } catch (e) {}
      done++;
      if (btn) btn.textContent = "상세정보 채우는 중… " + done + "/" + targets.length;
      scheduleRender();
      await sleep(180); // IP 차단 방지용 간격
    }
    enriching = false;
    if (btn) btn.textContent = "① 사용승인 채우기";
    flash(btn, "완료");
  }

  async function enrichOne(rec) {
    if (!IS_FIN) return; // basicInfo 엔드포인트는 fin.land 기준
    var id = rec.articleNo;
    var ret = rec.realEstateType;
    var tt = rec.tradeType;

    // 코드가 없으면 key API 로 조회
    if (!ret || !tt) {
      var keyRes = await pageFetch("https://fin.land.naver.com/front-api/v1/article/key?articleId=" + id);
      if (keyRes && keyRes.ok && keyRes.data) {
        var kret = deepFind(keyRes.data, function (k) { return /realEstateType/i.test(k); });
        var ktt = deepFind(keyRes.data, function (k) { return /tradeType/i.test(k); });
        if (kret) { ret = kret; rec.realEstateType = kret; }
        if (ktt) { tt = ktt; rec.tradeType = ktt; }
      }
    }

    var url = "https://fin.land.naver.com/front-api/v1/article/basicInfo?articleId=" + id;
    if (ret) url += "&realEstateType=" + encodeURIComponent(ret);
    if (tt) url += "&tradeType=" + encodeURIComponent(tt);
    var res = await pageFetch(url);
    if (res && res.ok && res.data) {
      var patch = extractDetailPatch(res.data, id);
      if (!patch) patch = { articleNo: id };
      if (!patch.useApproveYmd) {
        var uy = deepFind(res.data, function (k, v) {
          return /approve|aprv|사용승인|cmplt|준공/i.test(k) && extractYear(v);
        });
        if (uy) patch.useApproveYmd = uy;
      }
      // 좌표 캡처 — 상세 좌표가 더 정확하므로 목록 좌표를 덮어씀
      captureCoords(res.data, patch);
      var curRec = articles.get(id);
      if (curRec) {
        if (patch.lat != null) curRec.lat = patch.lat;
        if (patch.lng != null) curRec.lng = patch.lng;
      }
      // 상세에서 '가장 완전한' 주소(지번/번지 포함) 추출 — 중개사무소 주소는 제외
      var propAddr = extractPropertyAddress(res.data);
      dumpDetailAddress(id, res.data, propAddr); // 진단용
      var cur = articles.get(id);
      var curAddr = (cur && cur.address) || patch.address || "";
      // 새 주소가 번지(숫자)를 포함하고 기존보다 완전하면 교체
      if (propAddr && /\d/.test(propAddr) && propAddr.length >= curAddr.length) {
        patch.address = propAddr;
        if (cur) cur.address = propAddr; // 동→지번 업그레이드는 덮어쓰기
      } else if (!curAddr && propAddr) {
        patch.address = propAddr;
      }
      mergeRecord(patch);
    }
  }

  // 매물(부동산) 주소만 추출. 중개사(broker) 블록은 제외.
  function extractPropertyAddress(data) {
    // 중개사무소 주소가 섞이지 않도록 broker 계열 키를 먼저 제거
    var pruned = pruneKeys(data, /broker|agent|realtor|중개/i, 6);
    var addrObj = deepFindObject(pruned, function (k) { return /^address$|addressInfo|소재/i.test(k); });
    var s = addrObj ? bestAddressString(addrObj) : "";
    if (s && /\d/.test(s)) return s;
    var s2 = bestAddressString(pruned, 7);
    // 숫자(번지) 있는 쪽을 우선
    if (s2 && /\d/.test(s2) && (!s || !/\d/.test(s))) return s2;
    return s || s2;
  }
  // 특정 키(정규식)에 해당하는 하위 트리를 제거한 얕은 복제본
  function pruneKeys(obj, re, depth) {
    if (depth < 0 || !obj || typeof obj !== "object") return obj;
    if (Array.isArray(obj)) return obj.map(function (x) { return pruneKeys(x, re, depth - 1); });
    var out = {};
    Object.keys(obj).forEach(function (k) {
      if (re.test(k)) return; // broker 등 제거
      out[k] = pruneKeys(obj[k], re, depth - 1);
    });
    return out;
  }

  function captureCoords(data, patch) {
    var lat = deepFind(data, function (k, v) {
      return /^(latitude|lat|y)$/i.test(k) && typeof v !== "object" && Math.abs(toNum(v)) > 0 && Math.abs(toNum(v)) < 90;
    });
    var lng = deepFind(data, function (k, v) {
      return /^(longitude|lng|lon|x)$/i.test(k) && typeof v !== "object" && Math.abs(toNum(v)) > 90 && Math.abs(toNum(v)) < 180;
    });
    if (lat != null) patch.lat = toNum(lat);
    if (lng != null) patch.lng = toNum(lng);
  }

  // 진단: 목록 매물의 주소 객체 구조를 펼쳐 담음(① 없이도 확인 가능)
  var addrProbe = [];
  function probeArticleAddress(found) {
    for (var i = 0; i < found.length && addrProbe.length < 4; i++) {
      var o = found[i];
      var addr = o && o.address;
      if (!addr || typeof addr !== "object") continue;
      var inner = {};
      Object.keys(addr).slice(0, 40).forEach(function (k) {
        var v = addr[k];
        inner[k] = (v && typeof v === "object") ? JSON.stringify(v).slice(0, 80) : trimVal(v);
      });
      addrProbe.push({ articleNo: getArticleNo(o), addressKeys: Object.keys(addr), addressSample: inner, assembled: addressToString(addr) });
    }
  }

  // 진단: 상세 응답의 주소 객체/좌표를 몇 건 담아둠(🐞 진단복사에 포함)
  var detailDump = [];
  function dumpDetailAddress(id, data, propAddr) {
    if (detailDump.length >= 6) return;
    var addrObj = deepFindObject(data, function (k) { return /^address$/i.test(k); });
    var sample = {};
    if (addrObj) {
      Object.keys(addrObj).forEach(function (k) {
        var v = addrObj[k];
        sample[k] = (v && typeof v === "object") ? "{object}" : trimVal(v);
      });
    }
    detailDump.push({ articleNo: id, extracted: propAddr || "", addressObjectKeys: addrObj ? Object.keys(addrObj) : [], addressSample: sample });
  }

  /* ================= 표 렌더링 ================= */
  function scheduleRender() {
    if (renderTimer) return;
    renderTimer = setTimeout(function () {
      renderTimer = null;
      render();
    }, 250);
  }

  // 평수·평단가 기준 면적: 전용 → 공급 → 토지 순 (전용면적 기준)
  function areaBasisInfo(rec) {
    if (rec.area2 != null && rec.area2 > 0) return { area: rec.area2, label: "전용" };
    if (rec.area1 != null && rec.area1 > 0) return { area: rec.area1, label: "공급" };
    if (rec.landArea != null && rec.landArea > 0) return { area: rec.landArea, label: "토지" };
    return null;
  }
  function areaBasis(rec) {
    var b = areaBasisInfo(rec);
    return b ? b.area : null;
  }
  function pyeongOf(rec) {
    var a = areaBasis(rec);
    return a ? a / PYEONG : null;
  }

  // 평단가(만원/평): 매매=매매가, 전세=보증금 기준. 월세는 '평당월세'로 따로 계산.
  function pyeongPrice(rec) {
    var py = pyeongOf(rec);
    if (!py) return null;
    var name = rec.tradeTypeName || "";
    var amount = null;
    if (/매매/.test(name)) amount = rec.dealManwon;
    else if (/전세/.test(name)) amount = rec.depositManwon;
    else if (/월세|단기/.test(name)) return null;
    else amount = rec.dealManwon != null ? rec.dealManwon : rec.depositManwon;
    if (amount == null) return null;
    return Math.round(amount / py);
  }

  // 평당월세(만원/평·월): 월세 매물용
  function pyeongRent(rec) {
    var py = pyeongOf(rec);
    if (!py || rec.rentManwon == null) return null;
    return Math.round((rec.rentManwon / py) * 10) / 10;
  }

  // 단일 면적 표기: "150㎡ (45.4평)"
  function areaOne(m2) {
    var n = toNum(m2);
    if (n == null || n <= 0) return "";
    return n + "㎡ (" + (n / PYEONG).toFixed(1) + "평)";
  }

  function areaText(rec) {
    var parts = [];
    if (rec.area1 != null) parts.push(rec.area1);
    if (rec.area2 != null) parts.push(rec.area2);
    if (!parts.length && rec.landArea != null) parts.push(rec.landArea);
    if (!parts.length) return "";
    var b = areaBasisInfo(rec);
    var py = b ? (b.area / PYEONG).toFixed(1) : "";
    return parts.join("/") + "㎡" + (py ? " (" + b.label + " " + py + "평)" : "");
  }

  function priceCells(rec) {
    return {
      deal: rec.dealManwon != null ? formatManwon(rec.dealManwon) : "",
      deposit: rec.depositManwon != null ? formatManwon(rec.depositManwon) : "",
      rent: rec.rentManwon != null ? formatManwon(rec.rentManwon) : ""
    };
  }

  /* ---------- 정렬 (여러 축을 독립적으로 조합) ---------- */
  var tradeSort = "";    // "" | "매매" | "전세" | "월세"  → 해당 거래유형을 위로
  var floorSort = "";    // "" | "asc"(1층순) | "desc"(상층순)
  var approvalSort = ""; // "" | "new"(신축순=년차 적은순) | "old"(구축순=년차 많은순)
  var lastOrder = "floor"; // 층·년차 둘 다 켜졌을 때 어느 쪽을 우선할지(가장 최근에 누른 것)

  // 사용승인 경과년차(숫자). elapsedYear 우선, 없으면 연도로 계산. 불명은 null.
  function approvalYears(rec) {
    if (rec.elapsedYear != null && !isNaN(rec.elapsedYear)) return rec.elapsedYear;
    var yr = extractYear(rec.useApproveYmd);
    if (yr) {
      var e = new Date().getFullYear() - parseInt(yr, 10) + 1;
      return e < 1 ? 1 : e;
    }
    return null;
  }

  // 층 정보 "3/4", "B1/4", "1/8" 등에서 해당 층 숫자 추출 (지하는 음수, 불명은 null)
  function floorNum(rec) {
    var f = String(rec.floor == null ? "" : rec.floor).trim();
    if (!f || f === "-") return null;
    var head = f.split("/")[0].trim();
    if (/^B/i.test(head) || /지하/.test(head)) {
      var b = parseInt(head.replace(/\D/g, ""), 10);
      return isNaN(b) ? -1 : -b;
    }
    if (/^반지/.test(head)) return 0;
    var n = parseInt(head, 10);
    return isNaN(n) ? null : n;
  }

  function confirmCmp(a, b) {
    return String(b.confirmYmd || "").localeCompare(String(a.confirmYmd || ""));
  }

  // 층 비교: floorSort 미선택이거나 동률이면 0. 층 불명은 항상 아래로.
  function cmpFloor(a, b) {
    if (!floorSort) return 0;
    var fa = floorNum(a), fb = floorNum(b);
    var na = fa == null, nb = fb == null;
    if (na && nb) return 0;
    if (na) return 1;
    if (nb) return -1;
    if (fa !== fb) return floorSort === "asc" ? fa - fb : fb - fa;
    return 0;
  }
  // 년차 비교: approvalSort 미선택이거나 동률이면 0. 년차 불명은 항상 아래로.
  function cmpApproval(a, b) {
    if (!approvalSort) return 0;
    var ya = approvalYears(a), yb = approvalYears(b);
    var ea = ya == null, eb = yb == null;
    if (ea && eb) return 0;
    if (ea) return 1;
    if (eb) return -1;
    if (ya !== yb) return approvalSort === "new" ? ya - yb : yb - ya;
    return 0;
  }

  function applySort(rows) {
    // 층·년차 중 가장 최근에 누른 쪽을 우선 적용
    var first = lastOrder === "approval" ? cmpApproval : cmpFloor;
    var second = lastOrder === "approval" ? cmpFloor : cmpApproval;
    rows.sort(function (a, b) {
      // 1순위: 거래유형 (선택 시 해당 유형을 위로 묶음)
      if (tradeSort) {
        var am = (a.tradeTypeName || "").indexOf(tradeSort) >= 0 ? 0 : 1;
        var bm = (b.tradeTypeName || "").indexOf(tradeSort) >= 0 ? 0 : 1;
        if (am !== bm) return am - bm;
      }
      // 2·3순위: 층/년차 (최근에 누른 쪽 우선)
      var r = first(a, b); if (r) return r;
      r = second(a, b); if (r) return r;
      // 마지막: 확인일 최신순
      return confirmCmp(a, b);
    });
    return rows;
  }

  /* ---------- 동일 매물(여러 중개사가 올린 같은 물건) 판정 ---------- */
  var dedupeOn = true;
  // 확인일 없는 매물(지도 마커/클러스터/초기 embedded 조각) 숨기기
  var hideIncomplete = true;
  function isRealListing(rec) {
    // 실제 매물 목록 카드는 확인일(articleConfirmDate)을 가짐. 없으면 조각으로 간주.
    return !!rec.confirmYmd;
  }

  // 다른 지역(이전에 보던 지역·추천 단지 등)의 매물 숨기기
  var regionFilterOn = true;
  var dominantRegion = "";
  var hiddenRegionCount = 0;
  // 주소의 시·구 단위 키. 예: "경기도 시흥시 장곡동" -> "경기도 시흥시"
  function regionKey(rec) {
    if (!rec.address) return "";
    var toks = String(rec.address).trim().split(/\s+/);
    if (toks.length >= 2) return toks[0] + " " + toks[1];
    return toks[0] || "";
  }

  // 다른 종류(검색과 무관하게 끼어든 매물) 숨기기
  var typeFilterOn = true;
  var dominantType = "";
  var dominantTypeName = "";
  var hiddenTypeCount = 0;
  function typeKey(rec) {
    return rec.realEstateType || rec.realEstateTypeName || "";
  }

  function normTxt(v) {
    return String(v == null ? "" : v).replace(/\s+/g, "").toLowerCase();
  }

  // 같은 물건이면 같은 키. 정보가 부족하면 null(=중복 판정하지 않음).
  function dedupeKey(rec) {
    var hasArea = rec.area1 != null || rec.area2 != null || rec.landArea != null;
    var hasPlace = !!(rec.address || rec.name);
    // 면적도 장소정보도 없으면 같은 물건인지 판단할 근거가 없음 → 합치지 않음
    if (!hasArea && !hasPlace) return null;
    return [
      normTxt(rec.realEstateType),
      normTxt(rec.tradeTypeName),
      rec.dealManwon == null ? "" : rec.dealManwon,
      rec.depositManwon == null ? "" : rec.depositManwon,
      rec.rentManwon == null ? "" : rec.rentManwon,
      rec.area1 == null ? "" : rec.area1,
      rec.area2 == null ? "" : rec.area2,
      rec.landArea == null ? "" : rec.landArea,
      normTxt(rec.floor),
      normTxt(rec.address || rec.name)
    ].join("|");
  }

  // 정보가 더 충실한 레코드를 대표로 남기기 위한 점수
  function completeness(r) {
    var s = 0;
    if (r.address) s += 3;
    if (r.useApproveYmd) s += 2;
    if (r.name) s += 1;
    if (r.floor) s += 1;
    if (r.confirmYmd) s += 1;
    if (r.area1 != null) s += 1;
    if (r.area2 != null) s += 1;
    if (r.direction) s += 1;
    return s;
  }

  var hiddenIncompleteCount = 0;
  function toRows() {
    // 디스코 실거래: 필터·중복병합 없이 전부 표시(articleNo가 이미 유일키), 거래일순 정렬만
    if (IS_DISCO) {
      hiddenIncompleteCount = 0; hiddenRegionCount = 0; hiddenTypeCount = 0; dominantRegion = ""; dominantTypeName = "";
      var drows = [];
      articles.forEach(function (rec) { rec._dup = 1; drows.push(rec); });
      applySort(drows);
      return drows;
    }
    var all = [];
    hiddenIncompleteCount = 0;
    articles.forEach(function (rec) {
      rec._dup = 1;
      if (hideIncomplete && !isRealListing(rec)) { hiddenIncompleteCount++; return; }
      all.push(rec);
    });

    var rows;
    if (!dedupeOn) {
      rows = all;
    } else {
      var groups = new Map();
      rows = [];
      for (var i = 0; i < all.length; i++) {
        var r = all[i];
        var k = dedupeKey(r);
        if (k == null) { rows.push(r); continue; }
        var g = groups.get(k);
        if (!g) {
          groups.set(k, r);
          rows.push(r);
        } else {
          var n = (g._dup || 1) + 1;
          if (completeness(r) > completeness(g)) {
            // 더 충실한 쪽을 대표로 교체
            var idx = rows.indexOf(g);
            r._dup = n;
            if (idx >= 0) rows[idx] = r;
            groups.set(k, r);
          } else {
            g._dup = n;
          }
        }
      }
    }
    // 가장 많은 지역(시·구)을 자동 감지 → 다른 지역 매물 숨기기
    hiddenRegionCount = 0;
    dominantRegion = "";
    var counts = {};
    for (var m = 0; m < rows.length; m++) {
      var rk = regionKey(rows[m]);
      if (!rk) continue;
      counts[rk] = (counts[rk] || 0) + 1;
      if (counts[rk] > (counts[dominantRegion] || 0)) dominantRegion = rk;
    }
    if (regionFilterOn && dominantRegion) {
      rows = rows.filter(function (r) {
        var rk = regionKey(r);
        if (rk && rk !== dominantRegion) { hiddenRegionCount++; return false; }
        return true; // 주소 없는 매물은 판단 불가 → 남김
      });
    }

    // 가장 많은 종류(아파트/토지/상가 등)를 자동 감지 → 다른 종류 매물 숨기기
    hiddenTypeCount = 0;
    dominantType = "";
    dominantTypeName = "";
    var tcounts = {};
    for (var p = 0; p < rows.length; p++) {
      var tk = typeKey(rows[p]);
      if (!tk) continue;
      tcounts[tk] = (tcounts[tk] || 0) + 1;
      if (tcounts[tk] > (tcounts[dominantType] || 0)) dominantType = tk;
    }
    dominantTypeName = dominantType ? retName(dominantType, "") : "";
    if (typeFilterOn && dominantType) {
      rows = rows.filter(function (r) {
        var tk = typeKey(r);
        if (tk && tk !== dominantType) { hiddenTypeCount++; return false; }
        return true;
      });
    }

    // 정렬 (기본: 확인일 최신순 / 거래유형 / 층)
    applySort(rows);
    return rows;
  }

  var COLUMNS = IS_DISCO ? [
    { key: "confirm", label: "거래일" },
    { key: "address", label: "주소" },
    { key: "name", label: "건물명" },
    { key: "realEstateTypeName", label: "종류" },
    { key: "useDistrict", label: "용도지역" },
    { key: "jimok", label: "지목" },
    { key: "deal", label: "실거래가" },
    { key: "pyeong", label: "평단가(만/평)" },
    { key: "landAreaTxt", label: "토지면적" },
    { key: "buildingAreaTxt", label: "건물면적" },
    { key: "useYear", label: "사용승인" },
    { key: "floor", label: "층수" }
  ] : [
    { key: "address", label: "주소" },
    { key: "realEstateTypeName", label: "종류" },
    { key: "tradeTypeName", label: "거래" },
    { key: "deal", label: "매매가" },
    { key: "deposit", label: "보증금" },
    { key: "rent", label: "월세" },
    { key: "pyeong", label: "평단가(만/전용평)" },
    { key: "pyeongRent", label: "평당월세(전용)" },
    { key: "area", label: "면적(㎡/평)" },
    { key: "useYear", label: "사용승인(년차)" },
    { key: "floor", label: "층" },
    { key: "name", label: "단지/건물" }
  ];

  function rowValues(rec) {
    var pc = priceCells(rec);
    // 디스코가 직접 준 평단가(site_py_price)가 있으면 그걸 우선 사용, 없으면 계산
    var pp = rec.discoPyeong != null ? Math.round(rec.discoPyeong) : pyeongPrice(rec);
    var pr = pyeongRent(rec);
    return {
      address: rec.address || "",
      realEstateTypeName: rec.realEstateTypeName || "",
      tradeTypeName: rec.tradeTypeName || "",
      deal: pc.deal,
      deposit: pc.deposit,
      rent: pc.rent,
      pyeong: pp != null ? pp.toLocaleString("ko-KR") : "",
      pyeongRent: pr != null ? String(pr) : "",
      area: areaText(rec),
      useYear: approvalText(rec),
      floor: rec.floor || "",
      confirm: formatYmd(rec.confirmYmd),
      name: rec.name || "",
      useDistrict: rec.useDistrict || "",
      jimok: rec.jimok || "",
      // 토지면적: 집합건물이 아니면 표시 / 건물면적: 집합=전용, 일반건물=연면적
      landAreaTxt: (!rec.ziphap && rec.landArea) ? areaOne(rec.landArea) : "",
      buildingAreaTxt: rec.ziphap ? (rec.area2 ? areaOne(rec.area2) : "") : (rec.floorArea ? areaOne(rec.floorArea) : ""),
      dup: rec._dup > 1 ? rec._dup + "건" : "",
      articleNo: rec.articleNo
    };
  }

  // 사용승인 → 년차 표시 (네이버 방식). 예: "19년차 (2008)"
  function approvalText(rec) {
    var el = rec.elapsedYear;
    var yr = extractYear(rec.useApproveYmd);
    if (el == null && yr) {
      el = new Date().getFullYear() - parseInt(yr, 10) + 1; // 네이버는 준공연도를 1년차로 셈
      if (el < 1) el = 1;
    }
    if (el == null && !yr) return "";
    if (el == null) return yr ? "(" + yr + ")" : "";
    return el + "년차" + (yr ? " (" + yr + ")" : "");
  }

  // 거래유형 → 행 색상 클래스
  function tradeClass(name) {
    name = name || "";
    if (/매매/.test(name)) return "nre-t-deal";
    if (/전세/.test(name)) return "nre-t-jeonse";
    if (/단기/.test(name)) return "nre-t-short";
    if (/월세/.test(name)) return "nre-t-rent";
    return "";
  }

  function formatYmd(v) {
    if (!v) return "";
    var s = String(v);
    var m = s.match(/(\d{4})(\d{2})(\d{2})/);
    if (m) return m[1] + "." + m[2] + "." + m[3];
    return s;
  }

  /* ---------- 패널 DOM ---------- */
  var panel = null, tbodyEl = null, fab = null, countEl = null;

  // 헤더를 잡고 드래그하면 패널을 옮길 수 있게
  function makePanelDraggable(panelEl, handleEl) {
    if (!handleEl) return;
    var dragging = false, sx = 0, sy = 0, startLeft = 0, startTop = 0;
    handleEl.style.cursor = "move";
    handleEl.addEventListener("pointerdown", function (e) {
      // 버튼·체크박스·링크 등 조작 요소 위에서는 드래그 안 함
      if (e.target.closest("button, input, label, a, select, .nre-sbtn")) return;
      if (e.button != null && e.button !== 0) return;
      var r = panelEl.getBoundingClientRect();
      panelEl.style.left = r.left + "px";
      panelEl.style.top = r.top + "px";
      panelEl.style.right = "auto";
      panelEl.style.bottom = "auto";
      startLeft = r.left; startTop = r.top; sx = e.clientX; sy = e.clientY;
      dragging = true;
      try { handleEl.setPointerCapture(e.pointerId); } catch (ex) {}
      e.preventDefault();
    });
    handleEl.addEventListener("pointermove", function (e) {
      if (!dragging) return;
      var nx = startLeft + (e.clientX - sx);
      var ny = startTop + (e.clientY - sy);
      var vw = window.innerWidth, vh = window.innerHeight;
      // 헤더가 화면 밖으로 완전히 나가지 않게 살짝 제한
      nx = Math.max(40 - panelEl.offsetWidth, Math.min(nx, vw - 60));
      ny = Math.max(0, Math.min(ny, vh - 34));
      panelEl.style.left = nx + "px";
      panelEl.style.top = ny + "px";
    });
    function end(e) {
      if (!dragging) return;
      dragging = false;
      try { handleEl.releasePointerCapture(e.pointerId); } catch (ex) {}
    }
    handleEl.addEventListener("pointerup", end);
    handleEl.addEventListener("pointercancel", end);
  }

  function ensureUI() {
    if (!document.body) { setTimeout(ensureUI, 100); return; }
    if (fab) return;

    var appTitle = IS_DISCO ? "🧾 디스코 실거래 수집" : "📋 네이버부동산 매물정리";
    var fabTxt = IS_DISCO ? "실거래" : "매물정리";

    fab = document.createElement("div");
    fab.id = "nre-fab";
    fab.innerHTML = '<span class="nre-fab-ico">▤</span><span class="nre-fab-txt">' + fabTxt + '</span><span id="nre-fab-count">0</span>';
    fab.title = appTitle + " 표 열기";
    fab.addEventListener("click", togglePanel);
    document.body.appendChild(fab);
    countEl = fab.querySelector("#nre-fab-count");

    panel = document.createElement("div");
    panel.id = "nre-panel";
    panel.style.display = "none";
    panel.innerHTML =
      '<div id="nre-head">' +
        '<div class="nre-headrow">' +
          '<div class="nre-title">' + appTitle + ' <span id="nre-count2">0</span>건</div>' +
          '<div class="nre-headrow-right">' +
            '<div class="nre-user">👤 김명국 <span class="nre-user-mail">(mk1017sh@naver.com)</span></div>' +
            '<button id="nre-close" class="nre-btn nre-btn-close" title="닫기">✕</button>' +
          '</div>' +
        '</div>' +
        '<div class="nre-tools">' +
          '<label class="nre-chk" title="새 매물이 쌓일 때 좌표→지번을 자동으로 채웁니다(같은 좌표는 캐시 재사용)"><input type="checkbox" id="nre-autogeo" checked> 지번 자동</label>' +
          '<span class="nre-sort" title="거래유형으로 묶고, 층/년차 중 가장 최근에 누른 기준으로 정렬합니다. 같은 버튼 다시 누르면 해제">정렬:' +
            '<button class="nre-sbtn" data-sort="trade:매매">매매</button>' +
            '<button class="nre-sbtn" data-sort="trade:전세">전세</button>' +
            '<button class="nre-sbtn" data-sort="trade:월세">월세</button>' +
            '<span class="nre-sort-sep"></span>' +
            '<button class="nre-sbtn" data-sort="floor:asc">1층순</button>' +
            '<button class="nre-sbtn" data-sort="floor:desc">상층순</button>' +
            '<span class="nre-sort-sep"></span>' +
            '<button class="nre-sbtn" data-sort="appr:new">신축순</button>' +
            '<button class="nre-sbtn" data-sort="appr:old">구축순</button>' +
          '</span>' +
          '<button id="nre-csv" class="nre-btn">CSV</button>' +
          '<button id="nre-xls" class="nre-btn">엑셀</button>' +
          '<button id="nre-map" class="nre-btn" title="필지 맵핑 도구에서 바로 불러올 수 있는 좌표 포함 JSON으로 내보냅니다">🗺 지도</button>' +
          '<button id="nre-copy" class="nre-btn">복사</button>' +
          (IS_DISCO ? '<button id="nre-discodbg" class="nre-btn">🐞 진단복사</button>' : '') +
          '<button id="nre-clear" class="nre-btn">초기화</button>' +
        '</div>' +
      '</div>' +
      '<div id="nre-hint">' +
        (IS_DISCO
          ? '💡 디스코에서 <b>필지·건물을 클릭할 때마다</b> 그 물건의 용도지역·실거래가가 표에 쌓입니다(클릭한 것만). 금액 단위는 만원. '
          : '💡 매물 목록을 스크롤하면 자동으로 표에 쌓입니다. 평수·평단가는 <b>전용면적</b> 기준. ') +
        '거래유형 색: <span class="nre-lg nre-t-deal">매매</span> <span class="nre-lg nre-t-jeonse">전세</span> <span class="nre-lg nre-t-rent">월세</span> <span class="nre-lg nre-t-short">단기</span> ' +
        '<span id="nre-status"></span></div>' +
      '<div id="nre-tablewrap"><table id="nre-table"><thead></thead><tbody></tbody></table></div>';
    document.body.appendChild(panel);

    // 헤더
    var thead = panel.querySelector("thead");
    var trh = document.createElement("tr");
    trh.innerHTML = '<th>#</th>' + COLUMNS.map(function (c) { return "<th>" + c.label + "</th>"; }).join("") + "<th>링크</th>";
    thead.appendChild(trh);
    tbodyEl = panel.querySelector("tbody");
    tbodyEl.addEventListener("click", function (ev) {
      var t = ev.target;
      if (t && t.classList && t.classList.contains("nre-del")) {
        ev.preventDefault();
        var no = t.getAttribute("data-no");
        if (no) {
          // 해당 대표 매물 + 같은 물건으로 합쳐진 중복까지 함께 삭제
          var target = articles.get(no);
          var key = target ? dedupeKey(target) : null;
          articles.forEach(function (rec, k) {
            if (k === no || (dedupeOn && key != null && dedupeKey(rec) === key)) {
              dismissed.add(k);
              articles.delete(k);
            }
          });
          render();
        }
      }
    });

    panel.querySelector("#nre-close").addEventListener("click", togglePanel);
    makePanelDraggable(panel, panel.querySelector("#nre-head"));
    panel.querySelector("#nre-clear").addEventListener("click", function () {
      if (!confirm("표에 모인 매물을 모두 지울까요?\n\n지운 매물은 자동 재스캔으로 다시 담기지 않습니다.\n처음부터 다시 모으려면 페이지를 새로고침(F5)하세요.")) return;
      articles.forEach(function (_v, k) { dismissed.add(k); });
      articles.clear();
      render();
    });
    panel.querySelector("#nre-autogeo").addEventListener("change", function () {
      autoGeocode = this.checked;
      if (autoGeocode) { geoDisabled = false; enqueueGeocode(); }
    });
    panel.querySelector(".nre-sort").addEventListener("click", function (ev) {
      var b = ev.target;
      if (!b || !b.classList || !b.classList.contains("nre-sbtn")) return;
      var s = b.getAttribute("data-sort");
      if (s.indexOf("trade:") === 0) {
        var tv = s.slice(6);
        tradeSort = (tradeSort === tv) ? "" : tv; // 거래유형 세트: 같은 버튼 다시 누르면 해제
      } else if (s.indexOf("floor:") === 0) {
        var fv = s.slice(6);
        floorSort = (floorSort === fv) ? "" : fv; // 층 세트: 독립적으로 토글
        if (floorSort) lastOrder = "floor";       // 방금 누른 쪽을 우선
      } else if (s.indexOf("appr:") === 0) {
        var av = s.slice(5);
        approvalSort = (approvalSort === av) ? "" : av; // 년차 세트: 독립적으로 토글
        if (approvalSort) lastOrder = "approval";       // 방금 누른 쪽을 우선
      }
      render();
    });
    var discoDbgBtn = panel.querySelector("#nre-discodbg");
    if (discoDbgBtn) discoDbgBtn.addEventListener("click", function () {
      var report = { url: location.href, capturedResponses: stats.resp, collected: articles.size, discoResponses: discoDump };
      var text = JSON.stringify(report, null, 2);
      console.log("[실거래] 디스코 진단:", report);
      var b = this;
      try {
        navigator.clipboard.writeText(text).then(function () { flash(b, "복사됨! 붙여넣어 주세요"); }, function () { fallbackCopy(text, b); });
      } catch (e) { fallbackCopy(text, b); }
    });
    panel.querySelector("#nre-csv").addEventListener("click", exportCSV);
    panel.querySelector("#nre-xls").addEventListener("click", exportXLS);
    panel.querySelector("#nre-map").addEventListener("click", function () { exportMapJSON(panel.querySelector("#nre-map")); });
    panel.querySelector("#nre-copy").addEventListener("click", copyTable);
  }

  function togglePanel() {
    ensureUI();
    collectionEnabled = true; // 사용자가 처음 클릭한 시점부터 수집 시작
    var open = panel.style.display === "none";
    panel.style.display = open ? "flex" : "none";
    if (open) render();
  }

  function render() {
    ensureUI();
    var rows = toRows();
    lastCount = articles.size;
    if (countEl) countEl.textContent = rows.length;
    var c2 = panel && panel.querySelector("#nre-count2");
    if (c2) c2.textContent = rows.length;
    var st = panel && panel.querySelector("#nre-status");
    if (st) {
      var merged = articles.size - rows.length - hiddenIncompleteCount - hiddenRegionCount - hiddenTypeCount;
      st.textContent = "(가로챈 응답 " + stats.resp + "개 · 매물 " + rows.length + "건" +
        (merged > 0 ? " · 중복 " + merged + "건 합침" : "") +
        (hiddenIncompleteCount > 0 ? " · 확인일없음 " + hiddenIncompleteCount + "건 숨김" : "") +
        (hiddenRegionCount > 0 ? " · 타지역 " + hiddenRegionCount + "건 숨김" : "") +
        (hiddenTypeCount > 0 ? " · 타종류 " + hiddenTypeCount + "건 숨김" : "") +
        (geoDisabled ? " · ⚠지번 자동변환 중단(카카오맵 API 설정 필요)" : (geoStats.filled > 0 ? " · 지번 " + geoStats.filled + "곳" : "")) + ")";
      st.style.color = stats.resp === 0 ? "#c0392b" : (geoDisabled ? "#c0392b" : "#8a7a4a");
    }
    var rn = panel && panel.querySelector("#nre-region-name");
    if (rn) rn.textContent = dominantRegion ? " (" + dominantRegion + ")" : "";
    var tn = panel && panel.querySelector("#nre-type-name");
    if (tn) tn.textContent = dominantTypeName ? " (" + dominantTypeName + ")" : "";
    if (panel) {
      var sbtns = panel.querySelectorAll(".nre-sbtn");
      for (var b = 0; b < sbtns.length; b++) {
        var s = sbtns[b].getAttribute("data-sort");
        var on = (tradeSort && s === "trade:" + tradeSort) ||
          (floorSort && s === "floor:" + floorSort) ||
          (approvalSort && s === "appr:" + approvalSort);
        sbtns[b].classList.toggle("active", !!on);
      }
    }
    if (!tbodyEl || panel.style.display === "none") return;

    var html = "";
    for (var i = 0; i < rows.length; i++) {
      var v = rowValues(rows[i]);
      var rec = rows[i];
      var mapLink = (rec.lat != null && rec.lng != null)
        ? "<a href='https://map.kakao.com/link/map/" + encodeURIComponent((v.address || "위치") + "," + rec.lat + "," + rec.lng) +
          "' target='_blank' rel='noreferrer' title='카카오 지도에서 열기'>📍지도</a> "
        : "";
      var viewLink = IS_DISCO ? "" // 디스코 실거래는 개별 페이지가 없어 지도만 제공
        : "<a href='" + (IS_FIN
            ? "https://fin.land.naver.com/articles/" + v.articleNo
            : "https://new.land.naver.com/complexes?articleNo=" + v.articleNo) +
          "' target='_blank' rel='noreferrer'>보기</a> ";
      html += "<tr class='" + tradeClass(v.tradeTypeName) + "'>";
      html += "<td class='nre-idx'>" + (i + 1) + "</td>";
      for (var c = 0; c < COLUMNS.length; c++) {
        var val = v[COLUMNS[c].key] || "";
        html += "<td>" + escapeHtml(val) + "</td>";
      }
      html += "<td class='nre-actcell'>" + mapLink + viewLink +
        "<a href='#' class='nre-del' data-no='" + escapeHtml(v.articleNo) + "' title='이 항목 삭제'>✕</a></td>";
      html += "</tr>";
    }
    var emptyMsg = IS_DISCO ? "아직 수집된 실거래가 없습니다. 필지·건물을 클릭해 보세요." : "아직 수집된 매물이 없습니다. 목록을 스크롤해 보세요.";
    tbodyEl.innerHTML = html || "<tr><td colspan='" + (COLUMNS.length + 2) + "' class='nre-empty'>" + emptyMsg + "</td></tr>";
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, function (m) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m];
    });
  }

  function flash(btn, text) {
    if (!btn) return;
    var old = btn.textContent;
    btn.textContent = text;
    setTimeout(function () { btn.textContent = old; }, 1500);
  }

  /* ================= 내보내기 ================= */
  function buildMatrix() {
    var head = ["번호"].concat(COLUMNS.map(function (c) { return c.label; })).concat(["매물번호"]);
    var rows = toRows();
    var out = [head];
    for (var i = 0; i < rows.length; i++) {
      var v = rowValues(rows[i]);
      var line = [String(i + 1)];
      for (var c = 0; c < COLUMNS.length; c++) line.push(v[COLUMNS[c].key] || "");
      line.push(v.articleNo);
      out.push(line);
    }
    return out;
  }

  function download(filename, content, mime, noBom) {
    // CSV/XLS 는 엑셀 호환을 위해 BOM 을 붙이지만, JSON 은 BOM 이 있으면 JSON.parse 가 깨지므로 제외
    var blob = new Blob([noBom ? content : "﻿" + content], { type: mime + ";charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1000);
  }

  function stamp() {
    var d = new Date();
    function p(n) { return (n < 10 ? "0" : "") + n; }
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "_" + p(d.getHours()) + p(d.getMinutes());
  }

  function exportCSV() {
    var m = buildMatrix();
    var csv = m.map(function (row) {
      return row.map(function (cell) {
        var s = String(cell == null ? "" : cell);
        if (/[",\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"';
        return s;
      }).join(",");
    }).join("\r\n");
    download("네이버부동산_매물정리_" + stamp() + ".csv", csv, "text/csv");
  }

  function exportXLS() {
    // 엑셀이 여는 HTML 테이블(.xls)
    var m = buildMatrix();
    var thtml = "<table border='1'>" + m.map(function (row, ri) {
      var tag = ri === 0 ? "th" : "td";
      return "<tr>" + row.map(function (cell) {
        return "<" + tag + ">" + escapeHtml(cell) + "</" + tag + ">";
      }).join("") + "</tr>";
    }).join("") + "</table>";
    var html = '<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"></head><body>' + thtml + "</body></html>";
    download("네이버부동산_매물정리_" + stamp() + ".xls", html, "application/vnd.ms-excel");
  }

  // 필지 맵핑 도구용: 좌표 포함 JSON 내보내기 ("매물 불러오기"로 즉시 오버레이 생성)
  function exportMapJSON(btn) {
    if (IS_DISCO) { try { enqueueAddrGeocode(); } catch (e) {} } // 좌표 없는 디스코 항목 채우기 시도
    var rows = toRows();
    var items = [];
    var noCoord = 0;
    for (var i = 0; i < rows.length; i++) {
      var rec = rows[i];
      var v = rowValues(rec);
      if (rec.lat == null || rec.lng == null) noCoord++;
      items.push({
        source: IS_DISCO ? "disco" : "naver",
        articleNo: v.articleNo,
        address: v.address,
        lat: rec.lat != null ? rec.lat : null,
        lng: rec.lng != null ? rec.lng : null,
        realEstateTypeName: v.realEstateTypeName,
        tradeTypeName: v.tradeTypeName,
        tradeDate: IS_DISCO ? (v.confirm || "") : "", // 디스코 실거래 거래일만(네이버 confirm은 확인일이라 제외)
        deal: v.deal,
        deposit: v.deposit,
        rent: v.rent,
        pyeong: v.pyeong,
        pyeongRent: v.pyeongRent,
        area: v.area,
        useDistrict: v.useDistrict || "",
        jimok: v.jimok || "",
        landArea: v.landAreaTxt || "",
        buildingArea: v.buildingAreaTxt || "",
        useYear: v.useYear,
        floor: v.floor,
        name: v.name
      });
    }
    if (!items.length) { flash(btn, "내보낼 매물 없음"); return; }
    var payload = {
      type: "mapping-listings",
      version: 1,
      source: IS_DISCO ? "disco.re" : (IS_FIN ? "fin.land.naver.com" : "new.land.naver.com"),
      exportedAt: new Date().toISOString(),
      count: items.length,
      withoutCoords: noCoord,
      items: items
    };
    var fnPrefix = IS_DISCO ? "디스코실거래_지도용_" : "네이버매물_지도용_";
    download(fnPrefix + stamp() + ".json", JSON.stringify(payload, null, 2), "application/json", true);
    flash(btn, noCoord ? ("내보냄 · 좌표없음 " + noCoord + "건") : "지도용 내보냄!");
  }

  function copyTable() {
    var m = buildMatrix();
    var tsv = m.map(function (row) { return row.join("\t"); }).join("\n");
    var btn = panel.querySelector("#nre-copy");
    try {
      navigator.clipboard.writeText(tsv).then(function () { flash(btn, "복사됨!"); }, function () { fallbackCopy(tsv, btn); });
    } catch (e) { fallbackCopy(tsv, btn); }
  }

  function fallbackCopy(text, btn) {
    var ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); flash(btn, "복사됨!"); } catch (e) { flash(btn, "복사 실패"); }
    ta.remove();
  }

  /* ================= 시작 ================= */
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { ensureUI(); requestRescan(); });
  } else {
    ensureUI();
  }
  window.addEventListener("load", function () { setTimeout(requestRescan, 500); });

  // 서버렌더링 데이터 재스캔: 초반엔 자주, 이후엔 뜸하게
  var rescanCount = 0;
  var rescanTimer = setInterval(function () {
    requestRescan();
    rescanCount++;
    if (rescanCount > 15) {
      clearInterval(rescanTimer);
      rescanTimer = setInterval(requestRescan, 6000);
    }
  }, 2000);

  // 스크롤로 추가 매물 로드/페이지 이동 시 재스캔(쓰로틀)
  var scrollThrottle = 0;
  window.addEventListener("scroll", function () {
    var now = Date.now();
    if (now - scrollThrottle < 1200) return;
    scrollThrottle = now;
    requestRescan();
  }, true);

  requestRescan();
})();
