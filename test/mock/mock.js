// 練習ページ共通スクリプト（cas_auto_report の検証用）
// 業務システムの「対象者DB」「印字項目」「発行済みリスト」を模した最小実装。

// 対象者DB（存在しない宛名番号を渡すと検索0件になる）
var MOCK_DB = {
  "1234567890": "山田 太郎",
  "1234567891": "佐藤 花子",
  "1234567892": "鈴木 一郎"
};

// 印字項目（チェックボックス）
var MOCK_ITEMS = [
  { id: "chk-setai",    name: "世帯主" },
  { id: "chk-zokugara", name: "続柄" },
  { id: "chk-honseki",  name: "本籍" },
  { id: "chk-mynumber", name: "個人番号" }
];

var ISSUED_KEY = "mock_issued";
var FLIP_KEY   = "mock_flip";
var MULTI_KEY  = "mock_multi";

// 検証用スイッチ。ONにすると印字項目の初期状態を反転する（setcheckの実証用）
function mockFlipped() {
  try { return localStorage.getItem(FLIP_KEY) === "1"; } catch (e) { return false; }
}

// 宛名番号から初期チェック状態を決める（対象者ごとに初期状態がまちまちになる）
function mockInitialChecks(atena) {
  var n = 0;
  for (var i = 0; i < atena.length; i++) { n = (n * 31 + atena.charCodeAt(i)) % 16; }
  var flip = mockFlipped();
  var res = {};
  for (var j = 0; j < MOCK_ITEMS.length; j++) {
    var on = ((n >> j) & 1) === 1;
    res[MOCK_ITEMS[j].id] = flip ? !on : on;
  }
  return res;
}

// 検証用スイッチ。ONにすると氏名だけの一覧（roster.html）の検索結果を複数件にする
// （位置で押すとき「2件以上なら押さない」を確かめるため）
function mockMulti() {
  try { return localStorage.getItem(MULTI_KEY) === "1"; } catch (e) { return false; }
}

function mockLoadIssued() {
  try { return JSON.parse(localStorage.getItem(ISSUED_KEY) || "[]"); } catch (e) { return []; }
}

function mockSaveIssued(list) {
  try { localStorage.setItem(ISSUED_KEY, JSON.stringify(list)); } catch (e) {}
}

// 発行済みリストを描画する（帳票がサーバに溜まるのを模した代用）
function mockRenderIssued(box) {
  if (!box) { return; }
  var list = mockLoadIssued();
  box.innerHTML = "";
  for (var i = 0; i < list.length; i++) {
    var rec = list[i];
    var parts = [];
    for (var j = 0; j < MOCK_ITEMS.length; j++) {
      var it = MOCK_ITEMS[j];
      parts.push(it.name + ":" + (rec.items[it.id] ? "ON" : "OFF"));
    }
    var div = document.createElement("div");
    div.className = "issued-item";
    var pv = (rec.preview === undefined) ? "" : (" / プレビュー:" + (rec.preview ? "済" : "未"));
    div.textContent = (i + 1) + ". " + rec.atena + " " + rec.name + " / " + parts.join(" ") + pv;
    box.appendChild(div);
  }
}

// 検証用スイッチのチェックボックスを localStorage と同期させる
function mockInitFlip(el) {
  if (!el) { return; }
  el.checked = mockFlipped();
  el.addEventListener("change", function () {
    try { localStorage.setItem(FLIP_KEY, el.checked ? "1" : "0"); } catch (e) {}
  });
}

// 検索結果を複数件にするスイッチも同じように同期させる
function mockInitMulti(el) {
  if (!el) { return; }
  el.checked = mockMulti();
  el.addEventListener("change", function () {
    try { localStorage.setItem(MULTI_KEY, el.checked ? "1" : "0"); } catch (e) {}
  });
}
