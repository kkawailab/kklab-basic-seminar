// 出席確認の確認コード計算（出席確認ページと照合ページで共用）
// 注意: この関数群や時刻の書式を変更すると、すでに提出された確認コードの照合ができなくなります。

function fieldNormalize(s) {
  return s.normalize("NFKC").replace(/[\s　]+/g, " ").trim();
}

function codeKey() {
  var p = ["7hQt", "mZ", "26", "kkBS"];
  return p[3] + p[2] + "-" + p[1] + p[0];
}

async function computeCode(pageId, studentId, studentName, timeText) {
  var input = [
    pageId,
    fieldNormalize(studentId),
    fieldNormalize(studentName),
    timeText.trim(),
    codeKey()
  ].join("|");
  var buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(input));
  var hex = Array.from(new Uint8Array(buf))
    .map(function (b) { return b.toString(16).padStart(2, "0"); })
    .join("")
    .toUpperCase();
  return hex.slice(0, 4) + "-" + hex.slice(4, 8);
}

function formatJst(y, mo, d, h, mi, s) {
  var pad = function (n) { return String(n).padStart(2, "0"); };
  return y + "年" + mo + "月" + d + "日 " + pad(h) + ":" + pad(mi) + ":" + pad(s);
}

// サーバー時刻の取得。端末の時計・タイムゾーン設定には依存しない。
// 取得できなかった場合は null を返す（端末時計へのフォールバックはしない）。
async function fetchServerTime() {
  try {
    var r = await fetch("https://timeapi.io/api/Time/current/zone?timeZone=Asia/Tokyo", { cache: "no-store" });
    if (r.ok) {
      var d = await r.json();
      if (d && d.year) return formatJst(d.year, d.month, d.day, d.hour, d.minute, d.seconds);
    }
  } catch (e) { /* 次の候補へ */ }
  try {
    var r2 = await fetch("https://worldtimeapi.org/api/timezone/Asia/Tokyo", { cache: "no-store" });
    if (r2.ok) {
      var d2 = await r2.json();
      var m = String(d2.datetime).match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
      if (m) return formatJst(+m[1], +m[2], +m[3], +m[4], +m[5], +m[6]);
    }
  } catch (e) { /* 取得失敗 */ }
  return null;
}
