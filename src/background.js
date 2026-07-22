/*
 * background.js — 서비스 워커
 * 카카오 좌표→주소(coord2address) 역지오코딩을 대신 호출합니다.
 * (콘텐츠 스크립트에서 직접 부르면 CORS에 막히므로, host_permissions 를 가진 워커가 호출)
 */
chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
  if (msg && msg.type === "reverseGeocode") {
    reverseGeocode(msg.lat, msg.lng, msg.key).then(sendResponse);
    return true; // 비동기 응답
  }
  if (msg && msg.type === "forwardGeocode") {
    forwardGeocode(msg.address, msg.key).then(sendResponse);
    return true;
  }
});

// 주소 → 좌표 (카카오 주소검색). 디스코 실거래(좌표 없음)의 지번을 좌표로 변환.
async function forwardGeocode(address, key) {
  if (!address) return { ok: false, error: "no-address" };
  if (!key) return { ok: false, error: "no-key" };
  var url = "https://dapi.kakao.com/v2/local/search/address.json?query=" + encodeURIComponent(address);
  try {
    var res = await fetch(url, { headers: { Authorization: "KakaoAK " + key } });
    if (!res.ok) {
      var b = ""; try { b = (await res.text()).slice(0, 200); } catch (e) {}
      return { ok: false, status: res.status, body: b };
    }
    var data = await res.json();
    var doc = data && data.documents && data.documents[0];
    if (!doc) return { ok: true, lat: null, lng: null };
    return { ok: true, lat: parseFloat(doc.y), lng: parseFloat(doc.x) };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

async function reverseGeocode(lat, lng, key) {
  if (lat == null || lng == null) return { ok: false, error: "no-coords" };
  if (!key) return { ok: false, error: "no-key" };
  var url = "https://dapi.kakao.com/v2/local/geo/coord2address.json" +
    "?x=" + encodeURIComponent(lng) + "&y=" + encodeURIComponent(lat);
  try {
    var res = await fetch(url, { headers: { Authorization: "KakaoAK " + key } });
    if (!res.ok) {
      var body = "";
      try { body = (await res.text()).slice(0, 200); } catch (e) {}
      return { ok: false, status: res.status, body: body };
    }
    var data = await res.json();
    var doc = data && data.documents && data.documents[0];
    var jibun = doc && doc.address && doc.address.address_name;       // 지번주소
    var road = doc && doc.road_address && doc.road_address.address_name; // 도로명주소
    return { ok: true, jibun: jibun || "", road: road || "" };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}
