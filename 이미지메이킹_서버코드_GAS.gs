/*******************************************************************
 *  「이미지메이킹」 학습 앱 — 서버 코드 (Google Apps Script)
 *  영남이공대학교 모델테이너과
 *
 *  하는 일
 *   1) 학생이 '진도 제출하기'를 누르면 구글 시트에 한 줄로 기록한다.
 *      같은 학번이 다시 제출하면 그 줄을 최신 내용으로 갱신한다.
 *   2) 교수용 화면이 명단을 요청하면 시트 내용을 돌려준다.
 *
 *  설치 순서는 「설치·운영 매뉴얼」 문서를 참고하세요.
 *******************************************************************/

/* ── 설정 ────────────────────────────────────────────────────────
   TEACHER_CODE : 교수용 화면 코드. 학습 앱 HTML의 TEACHER_CODE와 같아야 한다.
   SHEET_NAME   : 기록이 쌓일 시트(탭) 이름. 없으면 자동으로 만들어진다.
   ---------------------------------------------------------------- */
var TEACHER_CODE = 'yn2026';
var SHEET_NAME   = '진도현황';

var HEADERS = ['제출시각', '이름', '학번', '진도(%)', '완료 장', '워크시트', '퀴즈 문항', '정답률(%)', '학생표시시각'];

/* ── 시트 준비 ─────────────────────────────────────────────────── */
function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#1F3864').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ── 응답 헬퍼 ─────────────────────────────────────────────────── */
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── 학생 제출 처리 ────────────────────────────────────────────── */
function submit_(d) {
  var name = String(d.n || '').trim();
  var sid  = String(d.sid || '').trim();
  if (!name || !sid) return json_({ ok: false, error: '이름과 학번이 필요합니다' });
  if (name.length > 30 || sid.length > 30) return json_({ ok: false, error: '입력이 너무 깁니다' });

  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return json_({ ok: false, error: '잠시 후 다시 시도해 주세요' }); }

  try {
    var sh = getSheet_();
    var now = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
    var row = [now, name, sid,
               Number(d.p) || 0, Number(d.ch) || 0, Number(d.ws) || 0,
               Number(d.qd) || 0, Number(d.q) || 0, String(d.t || '')];

    var last = sh.getLastRow();
    var found = 0;
    if (last > 1) {
      var sids = sh.getRange(2, 3, last - 1, 1).getValues();   // 학번 열
      for (var i = 0; i < sids.length; i++) {
        if (String(sids[i][0]).trim() === sid) { found = i + 2; break; }
      }
    }
    if (found) sh.getRange(found, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    return json_({ ok: true, updated: !!found });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ── 교수용 명단 조회 ──────────────────────────────────────────── */
function roster_(d) {
  if (String(d.code || '') !== TEACHER_CODE) return json_({ ok: false, error: '코드가 맞지 않습니다' });
  var sh = getSheet_();
  var last = sh.getLastRow();
  var rows = [];
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      if (!String(v[2] || '').trim()) continue;
      rows.push({
        t:  String(v[0] || ''),      // 서버 기록 시각
        n:  String(v[1] || ''),      // 이름
        sid:String(v[2] || ''),      // 학번
        p:  Number(v[3]) || 0,       // 진도
        ch: Number(v[4]) || 0,       // 완료 장
        ws: Number(v[5]) || 0,       // 워크시트 수
        qd: Number(v[6]) || 0,       // 퀴즈 문항
        q:  Number(v[7]) || 0        // 정답률
      });
    }
  }
  rows.sort(function (a, b) { return b.p - a.p; });
  return json_({ ok: true, rows: rows, count: rows.length });
}

/* ── 요청 진입점 ───────────────────────────────────────────────── */
function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '잘못된 요청입니다' }); }
  if (d.action === 'submit') return submit_(d);
  if (d.action === 'roster') return roster_(d);
  return json_({ ok: false, error: '알 수 없는 요청입니다' });
}

function doGet(e) {
  // 브라우저로 주소를 열었을 때 동작 확인용
  var p = (e && e.parameter) || {};
  if (p.action === 'roster') return roster_(p);
  return json_({ ok: true, message: '이미지메이킹 학습 앱 서버가 정상 동작 중입니다.' });
}

/* ── 설치 확인용 (에디터에서 직접 실행해 보세요) ────────────────── */
function 설치확인() {
  var sh = getSheet_();
  Logger.log('시트 준비 완료: ' + sh.getName() + ' / 현재 기록 ' + Math.max(0, sh.getLastRow() - 1) + '건');
}
