/**
 * 스몰 힐링 뽀모도로 - 백엔드 (Google Apps Script)
 * 이 스크립트가 들어있는 구글시트를 DB로 사용합니다.
 */

var SHEET_MISSIONS = 'Missions';
var SHEET_LOGS = 'Logs';
var SHEET_QUOTES = 'Quotes';
var SHEET_DIARY = 'Diary';
var SHEET_USERS = 'Users';

var MISSIONS_HEADERS = ['id', 'content', 'character', 'reward_message', 'is_active'];
var LOGS_HEADERS = ['timestamp', 'session_type', 'mission_id', 'completed', 'note', 'user_id'];
var QUOTES_HEADERS = ['id', 'quote', 'song_title', 'song_artist', 'is_active'];
var DIARY_HEADERS = ['log_id', 'date', 'category', 'tags', 'content', 'sentiment', 'energy_level', 'notes', 'user_id'];
var USERS_HEADERS = ['user_id', 'username', 'nickname', 'password_hash', 'salt', 'token', 'created_at',
  'recovery_question', 'recovery_answer_hash', 'last_login', 'privacy_agreed_at'];

// 회원마다 따로 생기는 탭: "기록_아이디", "일기_아이디"
var USER_LOGS_PREFIX = '기록_';
var USER_DIARY_PREFIX = '일기_';
var USER_LOGS_HEADERS = ['timestamp', 'session_type', 'mission_id', 'completed', 'note'];
var USER_DIARY_HEADERS = ['log_id', 'date', 'category', 'tags', 'content', 'sentiment', 'energy_level', 'notes'];

// 아이디/비밀번호 찾기용 질문 (화면의 선택 목록과 똑같아야 함)
var RECOVERY_QUESTIONS = [
  '어릴 적 별명은?',
  '처음 키운 반려동물 이름은?',
  '가장 기억에 남는 여행지는?',
  '가장 좋아하는 음식은?',
  '다녔던 초등학교 이름은?'
];

/**
 * 미션 완료 후 흘러나오는 힐링 클래식.
 * 작곡가가 세상을 떠난 지 70년이 넘어 곡 자체의 저작권이 끝난(퍼블릭 도메인) 곡들만 골랐고,
 * 유튜브에서도 "copyright free" 연주를 먼저 찾아서 화면 안에 바로 재생한다.
 */
var HEALING_CLASSICS = [
  { title: '달빛 (Clair de Lune)', composer: '드뷔시', query: 'Debussy Clair de Lune' },
  { title: '짐노페디 1번', composer: '사티', query: 'Satie Gymnopedie No 1' },
  { title: 'G선상의 아리아', composer: '바흐', query: 'Bach Air on the G String' },
  { title: '캐논', composer: '파헬벨', query: 'Pachelbel Canon in D' },
  { title: '녹턴 2번 (Op.9 No.2)', composer: '쇼팽', query: 'Chopin Nocturne Op 9 No 2' },
  { title: '트로이메라이 (꿈)', composer: '슈만', query: 'Schumann Traumerei' },
  { title: '백조', composer: '생상스', query: 'Saint-Saens The Swan' },
  { title: '타이스의 명상곡', composer: '마스네', query: 'Massenet Meditation Thais' },
  { title: '아베 마리아', composer: '슈베르트', query: 'Schubert Ave Maria instrumental' },
  { title: '사랑의 인사', composer: '엘가', query: 'Elgar Salut d Amour' }
];
// 검색이 실패할 때 대신 틀어주는 저작권 없는 클래식 연속 재생 영상 (배경음악과 같은 영상)
var FALLBACK_CLASSIC_VIDEO_ID = 'zJO0t7qqso4';

var GEMINI_MODEL = 'gemini-flash-lite-latest';

var SAMPLE_MISSIONS = [
  [1, '창밖을 10초간 바라보기', '모찌냥', '잘했어! 눈도 같이 쉬어줬네 💮', true],
  [2, '물 한 잔 마시기', '토끼모찌', '수분 보충 완료! 몸이 좋아할 거야 🥤', true],
  [3, '목과 어깨 가볍게 스트레칭하기', '병아리츠키', '으쌰으쌰! 몸이 한결 가벼워졌지? ✨', true],
  [4, '좋아하는 노래 한 소절 흥얼거리기', '모찌냥', '노래 한 곡에 기분이 사르르 풀리네~ 🎵', true],
  [5, '창문 열고 심호흡 세 번 하기', '토끼모찌', '상쾌한 공기 가득! 다시 힘이 나지? 🌿', true],
  [6, '오늘 고생한 나 자신을 칭찬해주기', '병아리츠키', '정말정말 잘하고 있어! 네가 자랑스러워 💛', true]
];

var SAMPLE_QUOTES = [
  [1, '지금 이 순간도 충분히 잘하고 있어요.', '밤편지', '아이유', true],
  [2, '잠깐 쉬어가도 괜찮아요, 그게 더 멀리 가는 방법이에요.', '너를 만나', '폴킴', true],
  [3, '오늘 하루도 애쓴 나에게 작은 박수를 보내요 👏', '좋은날', '아이유', true],
  [4, '완벽하지 않아도 돼요, 지금처럼만 하면 충분해요.', '오르막길', '정인 & 윤종신', true],
  [5, '숨 한 번 크게 쉬고, 다시 천천히 가볼까요?', '안녕', '정승환', true],
  [6, '쉬는 것도 일의 일부예요.', '산책', '잔나비', true]
];

/** Missions / Logs / Quotes 탭이 없으면 만들고, 헤더와 샘플 데이터를 채워둔다 */
function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var missions = ss.getSheetByName(SHEET_MISSIONS);
  if (!missions) {
    missions = ss.insertSheet(SHEET_MISSIONS);
    missions.appendRow(MISSIONS_HEADERS);
    SAMPLE_MISSIONS.forEach(function (row) {
      missions.appendRow(row);
    });
  }

  var logs = ss.getSheetByName(SHEET_LOGS);
  if (!logs) {
    logs = ss.insertSheet(SHEET_LOGS);
    logs.appendRow(LOGS_HEADERS);
  }

  var quotes = ss.getSheetByName(SHEET_QUOTES);
  if (!quotes) {
    quotes = ss.insertSheet(SHEET_QUOTES);
    quotes.appendRow(QUOTES_HEADERS);
    SAMPLE_QUOTES.forEach(function (row) {
      quotes.appendRow(row);
    });
  }

  var diary = ss.getSheetByName(SHEET_DIARY);
  if (!diary) {
    diary = ss.insertSheet(SHEET_DIARY);
    diary.appendRow(DIARY_HEADERS);
  }

  var users = ss.getSheetByName(SHEET_USERS);
  if (!users) {
    users = ss.insertSheet(SHEET_USERS);
    users.appendRow(USERS_HEADERS);
    styleHeader_(users);
  }
  // 나중에 추가된 칸(찾기 질문 등)이 예전 Users 탭에 없으면 붙여준다
  ['recovery_question', 'recovery_answer_hash', 'last_login', 'privacy_agreed_at'].forEach(function (h) {
    ensureColumn_(users, h);
  });

  // 예전부터 있던 Logs/Diary 탭에는 "누구의 기록인지" 칸이 없으니 맨 오른쪽에 붙여준다
  ensureColumn_(logs, 'user_id');
  ensureColumn_(diary, 'user_id');

  return { missions: missions, logs: logs, quotes: quotes, diary: diary, users: users };
}

/** 시트 1행(제목줄)에 해당 칸이 없으면 맨 오른쪽에 추가하고, 그 칸의 위치(0부터)를 반환 */
function ensureColumn_(sheet, name) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = headers.indexOf(name);
  if (idx !== -1) return idx;
  sheet.getRange(1, lastCol + 1).setValue(name);
  return lastCol;
}

/** 제목줄을 굵게 + 고정해서 시트를 보기 편하게 */
function styleHeader_(sheet) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#ffe6eb');
}

/**
 * 회원 전용 탭("기록_아이디" 또는 "일기_아이디")을 돌려준다.
 * 처음 만들 때는, 예전 공용 Logs/Diary 탭에 있던 그 사람의 기록을 그대로 옮겨 담는다.
 * (공용 탭은 지우지 않고 백업으로 남겨 둠)
 */
function getUserSheet_(kind, user) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var isLogs = kind === 'logs';
  var name = (isLogs ? USER_LOGS_PREFIX : USER_DIARY_PREFIX) + user.username;

  var sheet = ss.getSheetByName(name);
  if (sheet) return sheet;

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    sheet = ss.getSheetByName(name); // 기다리는 사이 다른 요청이 만들었을 수도 있음
    if (sheet) return sheet;

    var headers = isLogs ? USER_LOGS_HEADERS : USER_DIARY_HEADERS;
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    styleHeader_(sheet);
    sheet.setTabColor(isLogs ? '#bae6fd' : '#ffb5c2');

    // 공용 탭에서 이 사람(user_id) 기록만 골라 옮기기
    var shared = isLogs ? ensureSheets_().logs : ensureSheets_().diary;
    var data = shared.getDataRange().getValues();
    var sharedHeaders = data[0];
    var userIdx = sharedHeaders.indexOf('user_id');
    var colIdx = headers.map(function (h) { return sharedHeaders.indexOf(h); });
    var rows = data.slice(1)
      .filter(function (r) { return userIdx !== -1 && sameUser_(r[userIdx], user.userId); })
      .map(function (r) { return colIdx.map(function (i) { return i === -1 ? '' : r[i]; }); });
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    return sheet;
  } finally {
    lock.releaseLock();
  }
}

/** 제목줄 기준으로 한 줄을 만들어 맨 아래에 추가 (칸 순서가 바뀌어도 안전하게) */
function appendByHeaders_(sheet, obj) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  sheet.appendRow(headers.map(function (h) {
    return obj.hasOwnProperty(h) ? obj[h] : '';
  }));
}

/* ======================= 회원가입 / 로그인 ======================= */

function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '::' + password, Utilities.Charset.UTF_8);
  return bytes.map(function (b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function getUsers_() {
  var sheets = ensureSheets_();
  var data = sheets.users.getDataRange().getValues();
  var headers = data[0];
  return data.slice(1).map(function (r, i) {
    var u = { _row: i + 2 };
    headers.forEach(function (h, j) { u[h] = r[j]; });
    return u;
  }).filter(function (u) { return u.user_id; });
}

function publicUser_(u) {
  return {
    userId: u.user_id,
    username: String(u.username),
    nickname: u.nickname || u.username,
    token: u.token,
    hasRecovery: !!(u.recovery_question && u.recovery_answer_hash),
    hasConsent: !!u.privacy_agreed_at
  };
}

/** Users 탭의 특정 사람 줄에서 칸 하나를 바꾼다 */
function setUserCell_(u, header, value) {
  var users = ensureSheets_().users;
  var headers = users.getRange(1, 1, 1, users.getLastColumn()).getValues()[0];
  users.getRange(u._row, headers.indexOf(header) + 1).setValue(value);
}

function findUserByName_(username) {
  username = String(username || '').trim().toLowerCase();
  return getUsers_().filter(function (x) { return String(x.username).toLowerCase() === username; })[0];
}

/** 찾기 질문 답변은 띄어쓰기·대소문자 차이를 무시하고 비교 (예: "모찌 냥" = "모찌냥") */
function normalizeAnswer_(answer) {
  return String(answer || '').replace(/\s+/g, '').toLowerCase();
}

/** 같은 아이디로 틀린 시도를 너무 많이 하면 10분 동안 막는다 (비밀번호 마구 대입 방지) */
function tooManyTries_(key) {
  var cache = CacheService.getScriptCache();
  var n = parseInt(cache.get('try_' + key) || '0', 10);
  return n >= 5;
}
function countTry_(key) {
  var cache = CacheService.getScriptCache();
  var n = parseInt(cache.get('try_' + key) || '0', 10) + 1;
  cache.put('try_' + key, String(n), 600);
}
function clearTries_(key) {
  CacheService.getScriptCache().remove('try_' + key);
}

/** 로그인 토큰으로 사용자를 찾는다. 없으면 null */
function findUserByToken_(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var cached = cache.get('tok_' + token);
  if (cached) return JSON.parse(cached);

  var match = getUsers_().filter(function (u) { return u.token && u.token === token; })[0];
  if (!match) return null;
  var user = publicUser_(match);
  cache.put('tok_' + token, JSON.stringify(user), 21600);
  return user;
}

/** 로그인이 꼭 필요한 기능에서 사용: 토큰이 틀리면 에러를 던져서 화면이 로그인창으로 돌아가게 함 */
function requireUser_(token) {
  var user = findUserByToken_(token);
  if (!user) throw new Error('LOGIN_REQUIRED');
  return user;
}

/** 회원가입: 아이디/비밀번호/닉네임/찾기 질문을 받아 Users 탭에 저장하고 바로 로그인 상태로 만들어 준다 */
function signUp(username, password, nickname, question, answer, agreed) {
  if (agreed !== true) {
    return { ok: false, message: '기록 열람 안내에 동의해야 가입할 수 있어요.' };
  }
  username = String(username || '').trim();
  password = String(password || '');
  nickname = String(nickname || '').trim() || username;
  question = String(question || '');
  answer = normalizeAnswer_(answer);

  if (RECOVERY_QUESTIONS.indexOf(question) === -1 || !answer) {
    return { ok: false, message: '아이디·비밀번호를 잊었을 때를 위해 찾기 질문과 답변을 정해 주세요.' };
  }

  if (!/^[A-Za-z0-9가-힣_]{2,20}$/.test(username)) {
    return { ok: false, message: '아이디는 2~20자의 한글·영문·숫자로 만들어 주세요.' };
  }
  if (password.length < 4) {
    return { ok: false, message: '비밀번호는 4자 이상으로 정해 주세요.' };
  }
  if (nickname.length > 12) {
    return { ok: false, message: '닉네임은 12자 이내로 정해 주세요.' };
  }

  var user;
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheets = ensureSheets_();
    var users = getUsers_();
    var taken = users.some(function (u) { return String(u.username).toLowerCase() === username.toLowerCase(); });
    if (taken) {
      return { ok: false, message: '이미 사용 중인 아이디예요. 다른 아이디를 써 주세요.' };
    }

    var isFirstUser = users.length === 0;
    var salt = Utilities.getUuid();
    user = {
      user_id: 'U' + new Date().getTime(),
      username: username,
      nickname: nickname,
      password_hash: hashPassword_(password, salt),
      salt: salt,
      token: Utilities.getUuid(),
      created_at: new Date(),
      recovery_question: question,
      recovery_answer_hash: hashPassword_(answer, salt),
      last_login: new Date(),
      privacy_agreed_at: new Date()
    };
    appendByHeaders_(sheets.users, user);

    // 맨 처음 가입한 사람(앱 주인)에게는, 로그인 기능이 생기기 전에 쌓인 기록을 그대로 이어준다
    if (isFirstUser) {
      claimOrphanRows_(sheets.logs, user.user_id);
      claimOrphanRows_(sheets.diary, user.user_id);
    }
  } finally {
    lock.releaseLock();
  }

  // 회원 전용 탭("기록_아이디", "일기_아이디")을 바로 만들어 둔다
  var pub = publicUser_(user);
  getUserSheet_('logs', pub);
  getUserSheet_('diary', pub);
  return { ok: true, user: pub };
}

/** user_id 칸이 비어 있는 예전 기록들을 해당 사용자의 것으로 채운다 */
function claimOrphanRows_(sheet, userId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  var col = ensureColumn_(sheet, 'user_id') + 1;
  var range = sheet.getRange(2, col, lastRow - 1, 1);
  var values = range.getValues().map(function (r) { return [r[0] ? r[0] : userId]; });
  range.setValues(values);
}

/** 로그인: 맞으면 사용자 정보 + 토큰을 돌려준다 (토큰은 기기에 저장해서 다음부터 자동 로그인) */
function logIn(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  var tryKey = 'login_' + username.toLowerCase();
  if (tooManyTries_(tryKey)) {
    return { ok: false, message: '여러 번 틀려서 잠시 잠겼어요. 10분 뒤에 다시 시도하거나 "아이디·비밀번호 찾기"를 이용해 주세요.' };
  }

  var u = findUserByName_(username);
  if (!u || hashPassword_(password, u.salt) !== u.password_hash) {
    countTry_(tryKey);
    return { ok: false, message: '아이디 또는 비밀번호가 맞지 않아요.' };
  }
  clearTries_(tryKey);

  if (!u.token) {
    u.token = Utilities.getUuid();
    setUserCell_(u, 'token', u.token);
  }
  setUserCell_(u, 'last_login', new Date());
  return { ok: true, user: publicUser_(u) };
}

/* ======================= 아이디 / 비밀번호 찾기 ======================= */

/** 아이디 찾기: 닉네임 + 찾기 질문 + 답변이 맞는 계정의 아이디를 알려준다 */
function findUsername(nickname, question, answer) {
  nickname = String(nickname || '').trim();
  answer = normalizeAnswer_(answer);
  var tryKey = 'findid_' + nickname;
  if (!nickname || !answer) return { ok: false, message: '닉네임과 답변을 모두 입력해 주세요.' };
  if (tooManyTries_(tryKey)) return { ok: false, message: '여러 번 틀려서 잠시 잠겼어요. 10분 뒤에 다시 시도해 주세요.' };

  var found = getUsers_().filter(function (u) {
    return String(u.nickname) === nickname &&
      u.recovery_question === question &&
      u.recovery_answer_hash === hashPassword_(answer, u.salt);
  }).map(function (u) { return String(u.username); });

  if (found.length === 0) {
    countTry_(tryKey);
    return { ok: false, message: '입력하신 정보와 맞는 계정을 찾지 못했어요. 닉네임, 질문, 답변을 다시 확인해 주세요.' };
  }
  clearTries_(tryKey);
  return { ok: true, usernames: found };
}

/** 비밀번호 재설정: 아이디 + 찾기 질문 + 답변이 맞으면 새 비밀번호로 바꾼다 */
function resetPassword(username, question, answer, newPassword) {
  username = String(username || '').trim();
  answer = normalizeAnswer_(answer);
  newPassword = String(newPassword || '');
  var tryKey = 'reset_' + username.toLowerCase();

  if (!username || !answer) return { ok: false, message: '아이디와 답변을 모두 입력해 주세요.' };
  if (newPassword.length < 4) return { ok: false, message: '새 비밀번호는 4자 이상으로 정해 주세요.' };
  if (tooManyTries_(tryKey)) return { ok: false, message: '여러 번 틀려서 잠시 잠겼어요. 10분 뒤에 다시 시도해 주세요.' };

  var u = findUserByName_(username);
  if (!u || !u.recovery_answer_hash || u.recovery_question !== question ||
      u.recovery_answer_hash !== hashPassword_(answer, u.salt)) {
    countTry_(tryKey);
    return { ok: false, message: '아이디, 찾기 질문, 답변 중 맞지 않는 것이 있어요. 다시 확인해 주세요.' };
  }
  clearTries_(tryKey);
  clearTries_('login_' + username.toLowerCase());

  // 비밀번호를 바꾸면 다른 기기에 남아 있던 자동 로그인도 풀리도록 토큰도 새로 발급
  if (u.token) CacheService.getScriptCache().remove('tok_' + u.token);
  setUserCell_(u, 'password_hash', hashPassword_(newPassword, u.salt));
  setUserCell_(u, 'token', Utilities.getUuid());
  return { ok: true, message: '비밀번호를 새로 정했어요! 새 비밀번호로 로그인해 주세요 🌷' };
}

/** 찾기 질문이 없는 예전 계정이 로그인 후 질문을 등록할 때 사용 */
function setRecovery(token, question, answer) {
  var user = requireUser_(token);
  answer = normalizeAnswer_(answer);
  if (RECOVERY_QUESTIONS.indexOf(question) === -1 || !answer) {
    return { ok: false, message: '질문을 고르고 답변을 입력해 주세요.' };
  }
  var u = findUserByName_(user.username);
  setUserCell_(u, 'recovery_question', question);
  setUserCell_(u, 'recovery_answer_hash', hashPassword_(answer, u.salt));
  CacheService.getScriptCache().remove('tok_' + token);
  return { ok: true };
}

/** 동의 전에 가입한 예전 계정이 로그인 후 기록 열람 안내에 동의할 때 사용 (동의한 날짜를 남김) */
function agreePrivacy(token) {
  var user = requireUser_(token);
  setUserCell_(findUserByName_(user.username), 'privacy_agreed_at', new Date());
  CacheService.getScriptCache().remove('tok_' + token);
  return { ok: true };
}

/** 저장해 둔 토큰이 아직 유효한지 확인 (자동 로그인용) */
function whoAmI(token) {
  var user = findUserByToken_(token);
  return user ? { ok: true, user: user } : { ok: false };
}

/** 셀 값을 비교하기 좋게 문자열로 */
function sameUser_(cellValue, userId) {
  return String(cellValue) === String(userId);
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 활성화된(is_active=TRUE) 미션 목록을 배열로 반환 (화면과 JSON API 양쪽에서 재사용) */
function getActiveMissions() {
  var sheets = ensureSheets_();
  var data = sheets.missions.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);

  var idIdx = headers.indexOf('id');
  var contentIdx = headers.indexOf('content');
  var charIdx = headers.indexOf('character');
  var rewardIdx = headers.indexOf('reward_message');
  var activeIdx = headers.indexOf('is_active');

  return rows
    .filter(function (r) {
      var v = r[activeIdx];
      return v === true || String(v).toUpperCase() === 'TRUE';
    })
    .map(function (r) {
      return {
        id: r[idIdx],
        content: r[contentIdx],
        character: r[charIdx],
        reward_message: r[rewardIdx]
      };
    });
}

/** 활성화된(is_active=TRUE) 힐링 문구/노래 추천 목록을 배열로 반환 */
function getActiveQuotes() {
  var sheets = ensureSheets_();
  var data = sheets.quotes.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);

  var quoteIdx = headers.indexOf('quote');
  var songTitleIdx = headers.indexOf('song_title');
  var songArtistIdx = headers.indexOf('song_artist');
  var activeIdx = headers.indexOf('is_active');

  return rows
    .filter(function (r) {
      var v = r[activeIdx];
      return v === true || String(v).toUpperCase() === 'TRUE';
    })
    .map(function (r) {
      return {
        quote: r[quoteIdx],
        song_title: r[songTitleIdx],
        song_artist: r[songArtistIdx]
      };
    });
}

/**
 * 미션 완료 화면에서 틀어줄 힐링 클래식 1곡을 골라 { title, composer, videoId } 로 반환.
 * - 다른 사이트(이 앱) 안에서 재생이 허용된 영상만 찾는다(videoEmbeddable)
 * - 한 번 찾은 영상 번호는 저장해 두고 재사용 (유튜브 검색 무료 사용량 아끼기)
 */
function getHealingClassic() {
  var pick = HEALING_CLASSICS[Math.floor(Math.random() * HEALING_CLASSICS.length)];
  var props = PropertiesService.getScriptProperties();
  var cacheKey = 'classic_' + pick.query.replace(/[^A-Za-z0-9]/g, '_');

  var videoId = props.getProperty(cacheKey);
  if (!videoId) {
    videoId = searchEmbeddableVideo_(pick.query + ' copyright free classical music') ||
      searchEmbeddableVideo_(pick.query + ' classical music');
    if (videoId) props.setProperty(cacheKey, videoId);
  }

  return {
    title: pick.title,
    composer: pick.composer,
    videoId: videoId || FALLBACK_CLASSIC_VIDEO_ID
  };
}

function searchEmbeddableVideo_(query) {
  try {
    var result = YouTube.Search.list('snippet', {
      q: query,
      maxResults: 1,
      type: 'video',
      videoEmbeddable: 'true',
      safeSearch: 'strict'
    });
    if (result.items && result.items.length > 0) return result.items[0].id.videoId;
  } catch (err) {
    // 실패하면 null → 대체 영상 사용
  }
  return null;
}

/** 제목/가수로 유튜브를 검색해서 가장 위에 뜨는 영상의 링크를 반환 (실패하면 null) */
function getSongVideoUrl(title, artist) {
  try {
    var query = [title, artist].filter(function (v) { return !!v; }).join(' ');
    if (!query) return null;

    var result = YouTube.Search.list('snippet', { q: query, maxResults: 1, type: 'video' });
    if (result.items && result.items.length > 0) {
      return 'https://www.youtube.com/watch?v=' + result.items[0].id.videoId;
    }
  } catch (err) {
    // 실패하면 null 반환 (화면에서 검색 링크로 대신 보여줌)
  }
  return null;
}

/** 로그인한 사람의 "기록_아이디" 탭을 날짜별로 집계해서 오늘/이번주/전체 완료 수와 연속 사용일을 반환 */
function getStats(token) {
  var user = requireUser_(token);
  var data = getUserSheet_('logs', user).getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);

  var tsIdx = headers.indexOf('timestamp');
  var completedIdx = headers.indexOf('completed');
  var tz = Session.getScriptTimeZone();

  function dateKey(d) {
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }

  var countsByDate = {};
  var totalCount = 0;

  rows.forEach(function (r) {
    var completedVal = r[completedIdx];
    var isCompleted = completedVal === true || String(completedVal).toUpperCase() === 'TRUE';
    if (!isCompleted) return;

    var ts = r[tsIdx];
    if (!(ts instanceof Date)) return;

    var key = dateKey(ts);
    countsByDate[key] = (countsByDate[key] || 0) + 1;
    totalCount++;
  });

  var today = new Date();
  var todayKey = dateKey(today);
  var todayCount = countsByDate[todayKey] || 0;

  var weekCount = 0;
  var last7 = [];
  for (var i = 6; i >= 0; i--) {
    var d = new Date();
    d.setDate(d.getDate() - i);
    var key = dateKey(d);
    var count = countsByDate[key] || 0;
    weekCount += count;
    last7.push({ date: key, count: count });
  }

  var streak = 0;
  var cursor = new Date();
  if (!countsByDate[todayKey]) {
    cursor.setDate(cursor.getDate() - 1);
  }
  while (true) {
    var cursorKey = dateKey(cursor);
    if (countsByDate[cursorKey]) {
      streak++;
      cursor.setDate(cursor.getDate() - 1);
    } else {
      break;
    }
  }

  return {
    todayCount: todayCount,
    weekCount: weekCount,
    totalCount: totalCount,
    streak: streak,
    last7: last7
  };
}

/** 오늘의 일기 한 줄을 "일기_아이디" 탭에 저장 (log_id는 자동으로 다음 번호를 붙임) */
function addDiaryEntry(token, content, sentiment, energyLevel) {
  var user = requireUser_(token);
  var diarySheet = getUserSheet_('diary', user);
  content = (content || '').toString().trim();
  sentiment = (sentiment || '').toString().trim();
  energyLevel = (energyLevel !== undefined && energyLevel !== null) ? String(energyLevel).trim() : '';

  if (!content && !sentiment) {
    return { result: 'error', message: '내용이 비어있어요.' };
  }

  var entries = getDiaryEntries_(diarySheet);
  var nextNum = entries.length + 1;
  entries.forEach(function (e) {
    var m = String(e.log_id || '').match(/(\d+)$/);
    if (m) nextNum = Math.max(nextNum, parseInt(m[1], 10) + 1);
  });
  var logId = 'LOG_' + ('000' + nextNum).slice(-3);
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  appendByHeaders_(diarySheet, {
    log_id: logId, date: today, category: 'journal', tags: '', content: content,
    sentiment: sentiment, energy_level: energyLevel, notes: ''
  });
  return { result: 'success' };
}

/** 일기 탭(회원 전용 탭 또는 공용 Diary 탭)을 구조화된 배열로 반환 (오래된 순) */
function getDiaryEntries_(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  var tz = Session.getScriptTimeZone();

  var idIdx = headers.indexOf('log_id');
  var dateIdx = headers.indexOf('date');
  var categoryIdx = headers.indexOf('category');
  var tagsIdx = headers.indexOf('tags');
  var contentIdx = headers.indexOf('content');
  var sentimentIdx = headers.indexOf('sentiment');
  var energyIdx = headers.indexOf('energy_level');
  var notesIdx = headers.indexOf('notes');

  return rows
    .filter(function (r) {
      return r[contentIdx];
    })
    .map(function (r) {
      var rawDate = r[dateIdx];
      var date = rawDate instanceof Date
        ? Utilities.formatDate(rawDate, tz, 'yyyy-MM-dd')
        : String(rawDate);
      return {
        log_id: r[idIdx],
        date: date,
        category: r[categoryIdx],
        tags: r[tagsIdx],
        content: r[contentIdx],
        sentiment: r[sentimentIdx],
        energy_level: r[energyIdx],
        notes: r[notesIdx]
      };
    });
}

/** Gemini에 질문을 보내고 답변 텍스트를 반환 (서버 혼잡 시 몇 초 간격으로 자동 재시도) */
function callGemini_(apiKey, prompt) {
  // 복잡한 질문(여러 연도 비교 등)은 AI가 "생각"하는 데도 토큰을 쓰기 때문에,
  // 넉넉하게 잡아두지 않으면 생각만 하다가 실제 답변 글자가 하나도 안 나올 수 있다.
  var generationConfig = { maxOutputTokens: 4096 };
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent';

  var json = {};
  // 앱스 스크립트는 한 번 실행에 6분 제한이 있고, 요청 하나가 오래 붙잡힐 수도 있어서
  // 재시도 횟수/대기시간을 너무 늘리면 "최대 실행 시간 초과" 에러로 이어질 수 있다. 짧게 유지.
  var maxTries = 2;
  var waitMs = 1000;

  for (var attempt = 1; attempt <= maxTries; attempt++) {
    var response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: generationConfig
      }),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    try { json = JSON.parse(response.getContentText()); } catch (parseErr) { json = {}; }

    if (code === 200) break;

    var isOverloaded = code === 503 || code === 429 ||
      /overload|high demand|unavailable/i.test((json.error && json.error.message) || '');
    if (!isOverloaded || attempt === maxTries) {
      var apiMsg = (json.error && json.error.message) || ('AI 서버 응답 코드 ' + code);
      return { answer: null, error: isOverloaded ? '혼잡' : apiMsg };
    }

    Utilities.sleep(waitMs);
  }

  var answer = json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts &&
    json.candidates[0].content.parts[0] &&
    json.candidates[0].content.parts[0].text;

  return { answer: answer };
}

/** 로그인한 사람이 쓴 일기만 바탕으로 Gemini에게 질문하고 답변 + 참고한 기록을 반환 */
function askDiaryAI(token, question) {
  var user = requireUser_(token);
  return askDiaryAIFor_(user, question);
}

/** user가 있으면 그 사람의 "일기_아이디" 탭, 없으면(테스트용) 공용 Diary 탭을 읽는다 */
function askDiaryAIFor_(user, question) {
  var entries = getDiaryEntries_(user ? getUserSheet_('diary', user) : ensureSheets_().diary);

  if (entries.length === 0) {
    return { answer: '아직 쓴 일기가 없어서 답해드릴 수 없어요. 먼저 오늘의 기록을 남겨보세요!', sources: [] };
  }

  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { answer: 'AI 기능이 아직 설정되지 않았어요. (관리자에게 문의해주세요)', sources: [] };
  }

  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var range = findTimeRange_(question);
  var recent;

  if (range) {
    // "1년 전", "작년", "2025년 3월" 같은 질문이면 그 기간에 쓴 일기만 골라서 보낸다
    var inRange = entries.filter(function (e) {
      return e.date >= range.start && e.date <= range.end;
    });
    if (inRange.length === 0) {
      // 기록이 없으면 AI가 다른 날짜 이야기를 지어내지 않도록, AI에게 묻지 않고 바로 안내
      return {
        answer: '📭 ' + range.label + '의 기록은 찾아볼 수 없어요.\n' +
          '그래도 그때의 당신 역시, 분명 당신답게 한 걸음씩 걷고 있었을 거예요 🌿',
        sources: []
      };
    }
    recent = inRange.slice(-50);
  } else {
    // 질문과 관련 있는 기록을 먼저 추리고, 없으면 최근 기록으로 대체 (보낼 내용을 줄여서 더 빠르게 응답)
    var keywords = (question || '').split(/\s+/).filter(function (w) { return w.length > 1; });
    var matched = entries.filter(function (e) {
      return keywords.some(function (w) {
        return (e.content && e.content.indexOf(w) !== -1) || (e.tags && e.tags.indexOf(w) !== -1);
      });
    });
    recent = (matched.length > 0 ? matched : entries).slice(-50);
  }

  var diaryText = recent.map(function (e) {
    var tagPart = e.tags ? (' [태그: ' + e.tags + ']') : '';
    var moodPart = e.sentiment ? (' (감정: ' + e.sentiment + (e.energy_level ? ', 에너지: ' + e.energy_level : '') + ')') : '';
    return e.date + tagPart + moodPart + ': ' + e.content;
  }).join('\n');

  var rangeNote = range
    ? '이 질문은 ' + range.label + '(' + range.start + ' ~ ' + range.end + ')에 대한 것이고, 아래 기록은 그 기간에 쓴 것만 골랐어.\n'
    : '';

  var prompt =
    '오늘 날짜는 ' + todayStr + '이야.\n' + rangeNote +
    '아래는 사용자가 쓴 일기 기록이야.\n\n' + diaryText +
    '\n\n위 기록을 충분히 참고해서 다음 질문에 다정하고 공감하는 말투로 답해줘. ' +
    '답변 안에 참고한 날짜를 자연스럽게 언급해줘. ' +
    '기록에 없는 시기나 일은 절대 지어내지 말고, 질문한 시기의 기록이 없으면 ' +
    '"그때의 기록은 찾아볼 수 없지만, 그때의 당신도 당신답게 걷고 있었을 거예요" 같은 말로 따뜻하게 답해줘.\n\n질문: ' + question;

  try {
    var result = callGemini_(apiKey, prompt);
    var answer = result.answer;

    if (!answer) {
      var msg = result.error === '혼잡'
        ? 'AI 서버가 지금 많이 혼잡해서 몇 번 다시 시도했는데도 안 됐어요. 잠시 후 다시 물어봐주세요.'
        : '답변을 가져오는 데 문제가 생겼어요. 잠시 후 다시 시도해주세요.' + (result.error ? ' (' + result.error + ')' : '');
      return { answer: msg, sources: [] };
    }

    // 질문 속 단어가 포함된 기록을 우선 보여주고, 없으면 최근 기록을 보여줌 ("타임머신 기록")
    return { answer: answer, sources: recent.slice().reverse() };
  } catch (err) {
    return { answer: '답변을 가져오는 데 문제가 생겼어요: ' + String(err), sources: [] };
  }
}

/**
 * 질문 속 "언제"에 해당하는 말을 찾아 날짜 범위로 바꿔준다. 없으면 null.
 * 예) "1년 전" → 정확히 1년 전 날짜 앞뒤 한 달 / "작년" → 작년 1월 1일~12월 31일
 *     "3개월 전", "지난달" → 그 달 전체 / "2주 전", "지난주" → 그 주 / "어제", "5일 전" → 그 날
 *     "2025년", "2025년 3월" → 그 해/그 달 전체
 */
function findTimeRange_(question) {
  var q = String(question || '').replace(/일주일/g, '1주').replace(/이틀/g, '2일').replace(/사흘/g, '3일');
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  function fmt(d) { return Utilities.formatDate(d, tz, 'yyyy-MM-dd'); }
  function days(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
  function monthRange(y, m) { // m: 0~11
    return { start: fmt(new Date(y, m, 1)), end: fmt(new Date(y, m + 1, 0)) };
  }
  function yearRange(y) { return { start: y + '-01-01', end: y + '-12-31' }; }
  function withLabel(r, label) { r.label = label; return r; }

  // 1) "2025년", "2025년 3월" 처럼 연도를 직접 말한 경우
  var ym = q.match(/((?:19|20)\d{2})\s*년(?:\s*(\d{1,2})\s*월)?/);
  if (ym) {
    var y = parseInt(ym[1], 10);
    if (ym[2]) {
      var mm = parseInt(ym[2], 10);
      return withLabel(monthRange(y, mm - 1), y + '년 ' + mm + '월');
    }
    return withLabel(yearRange(y), y + '년');
  }

  var thisYear = now.getFullYear();
  if (/재작년/.test(q)) return withLabel(yearRange(thisYear - 2), '재작년(' + (thisYear - 2) + '년)');
  if (/작년|지난\s*해/.test(q)) return withLabel(yearRange(thisYear - 1), '작년(' + (thisYear - 1) + '년)');
  if (/지난\s*달/.test(q)) {
    var lm = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return withLabel(monthRange(lm.getFullYear(), lm.getMonth()), '지난달(' + (lm.getMonth() + 1) + '월)');
  }
  if (/지난\s*주/.test(q)) return withLabel({ start: fmt(days(now, -13)), end: fmt(days(now, -7)) }, '지난주');
  if (/그제|그저께/.test(q)) { var d2 = fmt(days(now, -2)); return withLabel({ start: d2, end: d2 }, '그저께(' + d2 + ')'); }
  if (/어제/.test(q)) { var d1 = fmt(days(now, -1)); return withLabel({ start: d1, end: d1 }, '어제(' + d1 + ')'); }

  // 2) "1년 전", "한 달 전", "3개월 전", "2주 전", "5일 전"
  var NUM = { '한': 1, '두': 2, '세': 3, '네': 4, '다섯': 5, '일': 1, '이': 2, '삼': 3, '사': 4, '오': 5 };
  var rel = q.match(/(\d+|다섯|한|두|세|네|일|이|삼|사|오)\s*(년|해|개월|달|주|일)\s*전/);
  if (rel) {
    var n = /^\d+$/.test(rel[1]) ? parseInt(rel[1], 10) : NUM[rel[1]];
    var unit = rel[2];
    if (unit === '년' || unit === '해') {
      var center = new Date(now.getFullYear() - n, now.getMonth(), now.getDate());
      return withLabel({ start: fmt(days(center, -30)), end: fmt(days(center, 30)) },
        n + '년 전(' + center.getFullYear() + '년 ' + (center.getMonth() + 1) + '월 무렵)');
    }
    if (unit === '개월' || unit === '달') {
      var t = new Date(now.getFullYear(), now.getMonth() - n, 1);
      return withLabel(monthRange(t.getFullYear(), t.getMonth()),
        n + '개월 전(' + t.getFullYear() + '년 ' + (t.getMonth() + 1) + '월)');
    }
    if (unit === '주') {
      var wc = days(now, -7 * n);
      return withLabel({ start: fmt(days(wc, -3)), end: fmt(days(wc, 3)) }, n + '주 전');
    }
    var dd = fmt(days(now, -n));
    return withLabel({ start: dd, end: dd }, n + '일 전(' + dd + ')');
  }

  return null;
}

/** [테스트용] 질문에서 날짜 범위를 잘 찾아내는지 확인 */
function 테스트_날짜범위_찾기() {
  ['나는 1년 전에 어땠어?', '작년에 나 어땠어?', '3개월 전 기분은?', '지난주에 뭐 했지?',
   '2025년 3월에 나는?', '2019년의 나는?', '요즘 내 기분 어때?'].forEach(function (q) {
    Logger.log(q + ' → ' + JSON.stringify(findTimeRange_(q)));
  });
}

/** [진단용] 지금 스크립트 속성에 어떤 이름들이 등록되어 있는지 확인 */
function 진단_스크립트속성_확인하기() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var keys = Object.keys(props);
  Logger.log('등록된 속성 이름들: ' + (keys.length > 0 ? keys.join(', ') : '(하나도 없음)'));
}

/**
 * AI 일기장 기능을 처음 한 번 허용받기 위한 테스트 함수.
 * 1) 스크립트 속성에 GEMINI_API_KEY를 먼저 등록한 뒤
 * 2) Apps Script 편집기에서 이 함수를 선택하고 ▶ 실행 버튼을 한 번만 눌러주세요.
 */
function 처음한번_실행해서_AI권한허용하기() {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  Logger.log('API 키 등록 여부: ' + (apiKey ? '등록됨' : '아직 등록 안 됨'));
  var result = askDiaryAIFor_(null, '테스트 질문이에요');
  Logger.log('결과: ' + JSON.stringify(result));
}

/** [테스트용] 유튜브에서 힐링 클래식 영상을 잘 찾아오는지 확인 */
function 테스트_힐링클래식_찾기() {
  Logger.log(JSON.stringify(getHealingClassic()));
}

/** 로그인한 사람의 완료 기록 1건을 "기록_아이디" 탭에 저장 */
function logCompletion(token, payload) {
  var user = requireUser_(token);
  return saveLog_(payload, user);
}

/** 완료 기록 저장. user가 없으면(외부 JSON API) 공용 Logs 탭에 저장 */
function saveLog_(payload, user) {
  payload = payload || {};
  var sheet = user ? getUserSheet_('logs', user) : ensureSheets_().logs;

  appendByHeaders_(sheet, {
    timestamp: new Date(),
    session_type: payload.session_type || '',
    mission_id: payload.mission_id || '',
    completed: payload.completed === false ? false : true,
    note: payload.note || ''
  });

  return { result: 'success' };
}

/**
 * GET 요청 처리.
 * - ?api=missions 로 호출하면 미션 목록을 JSON으로 반환 (doGet 명세 요구사항)
 * - 그 외에는 화면(Index.html)을 보여준다
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.api === 'missions') {
    return jsonOutput_({ missions: getActiveMissions() });
  }
  if (e && e.parameter && e.parameter.api === 'quotes') {
    return jsonOutput_({ quotes: getActiveQuotes() });
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('스몰 힐링 뽀모도로')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** POST 요청 처리: 완료 기록을 Logs 탭에 저장 (doPost 명세 요구사항) */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    return jsonOutput_(saveLog_(payload, null));
  } catch (err) {
    return jsonOutput_({ result: 'error', message: String(err) });
  }
}

/** Index.html 안에서 Style.html / ClientScript.html을 불러오기 위한 헬퍼 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
