/**
 * 文脈勝負｜Google スプレッドシート連携 API
 *
 * このファイルを管理用スプレッドシートの「拡張機能 → Apps Script」に貼り付け、
 * setupWorkbook を一度実行してください。その後、ウェブアプリとしてデプロイします。
 */

var SHEETS = {
  SETTINGS: '設定',
  STUDENTS: '生徒名簿',
  QUESTIONS: '問題一覧',
  ATTEMPTS: '実施記録',
  ACTIVE: '進行中',
  INDIVIDUAL: '個人検索',
  DASHBOARD: '全体ダッシュボード',
  PROGRESS: '実施状況',
  FOLLOW: '要フォロー生徒'
};

var LABELS = ['A', 'B', 'C', 'D', 'E', 'F'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('文脈勝負')
    .addItem('初期設定を行う', 'setupWorkbook')
    .addSeparator()
    .addItem('全体集計を更新', 'refreshDashboards')
    .addItem('個人検索を更新', 'refreshIndividualSearch')
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  if (e.range.getSheet().getName() === SHEETS.INDIVIDUAL && e.range.getA1Notation() === 'B3') {
    refreshIndividualSearch();
  }
}

function setupWorkbook() {
  var ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  ss.setSpreadsheetTimeZone('Asia/Tokyo');

  var active = ss.getSheetByName(SHEETS.ACTIVE);
  if (!active) active = ss.insertSheet(SHEETS.ACTIVE);
  if (active.getLastRow() === 0) {
    active.getRange(1, 1, 1, 9).setValues([[
      '実施ID', '実施トークン', '生徒ID', '問題ID', '実施日', '開始日時', '状態', '結果JSON', 'テスト実施'
    ]]);
  }
  active.setFrozenRows(1);
  active.hideSheet();

  var roster = ss.getSheetByName(SHEETS.STUDENTS);
  if (roster) {
    roster.getRange('A:A').setNumberFormat('@');
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['利用中', '利用停止'], true)
      .setAllowInvalid(false)
      .build();
    roster.getRange('D2:D500').setDataValidation(rule);
  }

  var attempts = ss.getSheetByName(SHEETS.ATTEMPTS);
  if (attempts) attempts.getRange('B:B').setNumberFormat('@');

  refreshDashboards();
  refreshIndividualSearch();
  onOpen();

  SpreadsheetApp.getUi().alert(
    '初期設定が完了しました。\n次に「デプロイ」からウェブアプリとして公開してください。'
  );
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'status';
    if (action !== 'status') throw appError_('未対応の操作です。', 'UNKNOWN_ACTION');
    return jsonOutput_({ ok: true, data: status_() });
  } catch (error) {
    return jsonOutput_(errorPayload_(error));
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var result;
    switch (body.action) {
      case 'verifyStudent':
        result = verifyStudent_(body.studentId);
        break;
      case 'startAttempt':
        result = startAttempt_(body.studentId);
        break;
      case 'submitAttempt':
        result = submitAttempt_(body);
        break;
      case 'refreshDashboards':
        result = refreshDashboardsAction_();
        break;
      default:
        throw appError_('未対応の操作です。', 'UNKNOWN_ACTION');
    }
    return jsonOutput_({ ok: true, data: result });
  } catch (error) {
    return jsonOutput_(errorPayload_(error));
  }
}

function status_() {
  var config = getConfig_();
  var today = tokyoDate_();
  return {
    today: today,
    available: Boolean(getQuestionForDate_(today)),
    period: { start: config.PERIOD_START, end: config.PERIOD_END },
    studentIdDigits: Number(config.STUDENT_ID_DIGITS),
    timing: timingConfig_(config)
  };
}

function verifyStudent_(rawStudentId) {
  var config = getConfig_();
  var studentId = validateStudentId_(rawStudentId, config);
  var testMode = studentId === String(config.TEST_STUDENT_ID);
  var student = getStudent_(studentId);
  if (!student) throw appError_('この生徒IDは登録されていません。先生に確認してください。', 'NOT_REGISTERED');
  if (!student.active) throw appError_('この生徒IDは現在利用できません。先生に確認してください。', 'INACTIVE');

  return {
    studentId: student.studentId,
    fullName: student.fullName,
    campus: student.campus,
    testMode: testMode
  };
}

function startAttempt_(rawStudentId) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    var config = getConfig_();
    var studentId = validateStudentId_(rawStudentId, config);
    var testMode = studentId === String(config.TEST_STUDENT_ID);
    var student = getStudent_(studentId);
    if (!student) throw appError_('この生徒IDは登録されていません。先生に確認してください。', 'NOT_REGISTERED');
    if (!student.active) throw appError_('この生徒IDは現在利用できません。先生に確認してください。', 'INACTIVE');

    var today = tokyoDate_();
    var question = getQuestionForDate_(today);
    if (testMode && !question) question = getQuestions_()[0] || null;
    if (!question) {
      throw appError_(
        '本日の問題はありません。実施期間は' + config.PERIOD_START + '～' + config.PERIOD_END + 'です。',
        'NOT_AVAILABLE'
      );
    }

    if (hasSubmitted_(studentId, question.questionId)) {
      throw appError_('本日の問題はすでに実施済みです。再挑戦はできません。', 'COMPLETED');
    }

    var existing = getActiveAttemptForStudent_(studentId, question.questionId);
    if (existing) return attemptResponse_(existing, question, student, testMode, true, config);

    var attempt = {
      attemptId: Utilities.getUuid(),
      attemptToken: Utilities.getUuid() + Utilities.getUuid(),
      studentId: studentId,
      questionId: question.questionId,
      practiceDate: today,
      startedAt: new Date().toISOString(),
      status: 'started',
      resultJson: '',
      testMode: testMode
    };
    appendActiveAttempt_(attempt);
    return attemptResponse_(attempt, question, student, testMode, false, config);
  } finally {
    lock.releaseLock();
  }
}

function submitAttempt_(body) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(20000);
  try {
    if (!body.attemptId || !body.attemptToken) throw appError_('提出情報が不足しています。', 'INVALID_SUBMISSION');
    var active = getActiveAttemptById_(String(body.attemptId));
    if (!active || active.attemptToken !== String(body.attemptToken)) {
      throw appError_('実施情報を確認できません。', 'ATTEMPT_NOT_FOUND');
    }
    if (active.status === 'submitted' && active.resultJson) {
      return { result: JSON.parse(active.resultJson), duplicate: true };
    }

    var config = getConfig_();
    var question = getQuestionById_(active.questionId);
    if (!question) throw appError_('問題データを確認できません。', 'QUESTION_NOT_FOUND');

    var answer = Array.isArray(body.answer) ? body.answer.map(String).filter(function (label) {
      return LABELS.indexOf(label) >= 0;
    }) : [];
    var complete = answer.length === 6 && unique_(answer).length === 6;
    var started = new Date(active.startedAt).getTime();
    var now = Date.now();
    var locked = body.lockedAt ? new Date(body.lockedAt).getTime() : now;
    var validLockedAt = isFinite(locked) && locked >= started && locked <= now + 5000;
    var elapsedSeconds = Math.max(0, Math.floor(((validLockedAt ? locked : now) - started) / 1000));
    var timeLimit = active.testMode ? Number(config.ANSWER_SECONDS) : Number(config.TOTAL_SECONDS);
    var timedOut = !complete || elapsedSeconds >= timeLimit;
    var correct = complete && !timedOut && answer.join(',') === question.correctOrder.join(',');
    var feedbackRule = correct ? null : (
      firstBrokenRule_(question.feedbackRules, answer) || firstFeedbackRule_(question.feedbackRules)
    );
    var result = {
      correct: correct,
      timedOut: timedOut,
      answer: answer,
      correctOrder: question.correctOrder,
      feedbackRule: feedbackRule ? {
        priority: feedbackRule.priority,
        explanation: feedbackRule.explanation,
        first: feedbackRule.first,
        second: feedbackRule.second
      } : null,
      review: feedbackRule ? buildReview_(question, feedbackRule) : null,
      elapsedSeconds: Math.min(elapsedSeconds, 600)
    };

    var student = getStudent_(active.studentId);
    if (!student) throw appError_('生徒情報を確認できません。', 'STUDENT_NOT_FOUND');
    appendSubmittedAttempt_(active, student, question, result);
    updateActiveResult_(active.rowNumber, result);
    return { result: result, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

function getConfig_() {
  var sheet = requiredSheet_(SHEETS.SETTINGS);
  var values = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 2).getDisplayValues();
  var config = {};
  values.forEach(function (row) {
    if (row[0]) config[row[0]] = row[1];
  });
  ['PERIOD_START', 'PERIOD_END', 'STUDENT_ID_DIGITS', 'TEST_STUDENT_ID', 'READING_SECONDS', 'ANSWER_SECONDS', 'TOTAL_SECONDS'].forEach(function (key) {
    if (config[key] === undefined || config[key] === '') throw appError_('設定シートの ' + key + ' を確認してください。', 'CONFIG_ERROR');
  });
  return config;
}

function getStudents_() {
  var sheet = requiredSheet_(SHEETS.STUDENTS);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getDisplayValues()
    .filter(function (row) { return /^\d{8}$/.test(row[0]); })
    .map(function (row) {
      return {
        studentId: row[0],
        fullName: row[1],
        campus: row[2],
        active: row[3] !== '利用停止' && row[3].toLowerCase() !== 'false'
      };
    });
}

function getStudent_(studentId) {
  return getStudents_().filter(function (student) { return student.studentId === studentId; })[0] || null;
}

function getQuestions_() {
  var sheet = requiredSheet_(SHEETS.QUESTIONS);
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 16).getValues();
  return values.filter(function (row) { return row[0] && row[1]; }).map(function (row) {
    var rules;
    try {
      rules = JSON.parse(String(row[10] || '[]'));
    } catch (error) {
      throw appError_('問題 ' + row[0] + ' の誤答解説ルールJSONを確認してください。', 'QUESTION_DATA_ERROR');
    }
    return {
      questionId: String(row[0]),
      releaseDate: formatDateValue_(row[1]),
      title: String(row[2]),
      paragraphs: LABELS.map(function (label, index) { return { label: label, text: String(row[index + 3]) }; }),
      correctOrder: String(row[9]).split(/\s*→\s*/).filter(Boolean),
      feedbackRules: rules,
      theme: String(row[11] || ''),
      skills: String(row[12] || ''),
      difficulty: String(row[13] || ''),
      status: String(row[15] || '')
    };
  }).sort(function (a, b) { return a.releaseDate.localeCompare(b.releaseDate); });
}

function getQuestionForDate_(date) {
  return getQuestions_().filter(function (question) { return question.releaseDate === date; })[0] || null;
}

function getQuestionById_(questionId) {
  return getQuestions_().filter(function (question) { return question.questionId === questionId; })[0] || null;
}

function attemptResponse_(attempt, question, student, testMode, resumed, config) {
  return {
    attemptId: attempt.attemptId,
    attemptToken: attempt.attemptToken,
    startedAt: attempt.startedAt,
    question: {
      questionId: question.questionId,
      releaseDate: question.releaseDate,
      title: question.title,
      paragraphs: question.paragraphs
    },
    timing: timingConfig_(config),
    testMode: Boolean(testMode || attempt.testMode),
    resumed: resumed,
    student: { fullName: student.fullName, campus: student.campus }
  };
}

function timingConfig_(config) {
  return {
    readingSeconds: Number(config.READING_SECONDS),
    answerSeconds: Number(config.ANSWER_SECONDS),
    totalSeconds: Number(config.TOTAL_SECONDS)
  };
}

function appendActiveAttempt_(attempt) {
  var sheet = requiredSheet_(SHEETS.ACTIVE);
  sheet.appendRow([
    attempt.attemptId, attempt.attemptToken, attempt.studentId, attempt.questionId,
    attempt.practiceDate, new Date(attempt.startedAt), attempt.status, attempt.resultJson, attempt.testMode
  ]);
}

function getActiveAttempts_() {
  var sheet = requiredSheet_(SHEETS.ACTIVE);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues().map(function (row, index) {
    return {
      rowNumber: index + 2,
      attemptId: String(row[0]),
      attemptToken: String(row[1]),
      studentId: String(row[2]),
      questionId: String(row[3]),
      practiceDate: formatDateValue_(row[4]),
      startedAt: row[5] instanceof Date ? row[5].toISOString() : String(row[5]),
      status: String(row[6]),
      resultJson: String(row[7] || ''),
      testMode: row[8] === true || String(row[8]).toLowerCase() === 'true'
    };
  });
}

function getActiveAttemptById_(attemptId) {
  return getActiveAttempts_().filter(function (attempt) { return attempt.attemptId === attemptId; })[0] || null;
}

function getActiveAttemptForStudent_(studentId, questionId) {
  return getActiveAttempts_().filter(function (attempt) {
    return attempt.studentId === studentId && attempt.questionId === questionId && attempt.status === 'started';
  })[0] || null;
}

function updateActiveResult_(rowNumber, result) {
  requiredSheet_(SHEETS.ACTIVE).getRange(rowNumber, 7, 1, 2).setValues([['submitted', JSON.stringify(result)]]);
}

function getSubmittedRows_() {
  var sheet = requiredSheet_(SHEETS.ATTEMPTS);
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 14).getValues().map(function (row) {
    return {
      submittedAt: row[0], studentId: String(row[1]), fullName: String(row[2]), campus: String(row[3]),
      questionId: String(row[4]), practiceDate: formatDateValue_(row[5]), answer: String(row[6]),
      correctOrder: String(row[7]), result: String(row[8]), timedOut: String(row[9]) === '時間切れ',
      elapsedSeconds: Number(row[10] || 0), feedbackPriority: row[11] === '' ? null : Number(row[11]),
      attemptId: String(row[12]), startedAt: row[13]
    };
  }).filter(function (row) { return row.studentId && row.questionId; });
}

function hasSubmitted_(studentId, questionId) {
  return getSubmittedRows_().some(function (row) {
    return row.studentId === studentId && row.questionId === questionId;
  });
}

function appendSubmittedAttempt_(active, student, question, result) {
  requiredSheet_(SHEETS.ATTEMPTS).appendRow([
    new Date(), active.studentId, student.fullName, student.campus, question.questionId,
    active.practiceDate, result.answer.join(' → '), result.correctOrder.join(' → '),
    result.correct ? '正解' : '不正解', result.timedOut ? '時間切れ' : '', result.elapsedSeconds,
    result.feedbackRule ? result.feedbackRule.priority : '', active.attemptId, new Date(active.startedAt)
  ]);
}

function firstBrokenRule_(rules, answer) {
  return rules.slice().sort(function (a, b) { return Number(a.priority) - Number(b.priority); }).filter(function (rule) {
    var first = answer.indexOf(rule.first);
    var second = answer.indexOf(rule.second);
    return first < 0 || second !== first + 1;
  })[0] || null;
}

function firstFeedbackRule_(rules) {
  return rules.slice().sort(function (a, b) {
    return Number(a.priority) - Number(b.priority);
  })[0] || null;
}

function buildReview_(question, rule) {
  var focus = [rule.first, rule.second];
  var highlights = {};
  focus.forEach(function (label) {
    var paragraph = question.paragraphs.filter(function (item) { return item.label === label; })[0];
    highlights[label] = paragraph ? [bestHighlight_(paragraph.text, rule.explanation)] : [];
  });
  return {
    focusLabels: focus,
    highlights: highlights,
    advice: '下線部を比べ、' + focus[1] + 'の内容が' + focus[0] + 'を受けて初めて意味が通るかを確かめましょう。後ろの段落にある指示語・接続語・言い換えから、直前に必要な内容を逆向きに探すのがコツです。'
  };
}

function bestHighlight_(text, explanation) {
  var quoted = [];
  var regex = /「([^「」]{3,})」/g;
  var match;
  while ((match = regex.exec(explanation)) !== null) quoted.push(match[1]);
  quoted.sort(function (a, b) { return b.length - a.length; });
  for (var i = 0; i < quoted.length; i += 1) if (text.indexOf(quoted[i]) >= 0) return quoted[i];

  var sentences = text.match(/[^。！？]+[。！？]?/g) || [text];
  var explanationSet = bigramSet_(explanation);
  var best = sentences[0] || text;
  var bestScore = -1;
  sentences.forEach(function (sentence) {
    var score = 0;
    Object.keys(bigramSet_(sentence)).forEach(function (gram) { if (explanationSet[gram]) score += 1; });
    if (score > bestScore) { best = sentence.trim(); bestScore = score; }
  });
  return best;
}

function bigramSet_(text) {
  var compact = String(text).replace(/[\s。、！？「」『』（）・]/g, '');
  var set = {};
  for (var i = 0; i < compact.length - 1; i += 1) set[compact.slice(i, i + 2)] = true;
  return set;
}

function getStudentStats_(studentId, questions, submittedRows, targetDate) {
  var today = targetDate || tokyoDate_();
  var questionSource = questions || getQuestions_();
  var attemptSource = submittedRows || getSubmittedRows_();
  var releasedQuestions = questionSource.filter(function (question) {
    return question.releaseDate <= today;
  });
  var releasedIdSet = {};
  releasedQuestions.forEach(function (question) {
    releasedIdSet[question.questionId] = true;
  });

  var allRows = attemptSource.filter(function (row) {
    return row.studentId === studentId;
  });
  var releasedRows = allRows.filter(function (row) {
    return releasedIdSet[row.questionId];
  });
  var completedIds = unique_(releasedRows.map(function (row) {
    return row.questionId;
  }));
  var completedSet = {};
  completedIds.forEach(function (questionId) {
    completedSet[questionId] = true;
  });

  var correctSet = {};
  releasedRows.forEach(function (row) {
    if (row.result === '正解') correctSet[row.questionId] = true;
  });

  // 要フォローは「今日」を含めず、完了期限を過ぎた問題の連続未実施で判定する。
  var consecutiveMissedDays = 0;
  var completedDeadlineQuestions = releasedQuestions.filter(function (question) {
    return question.releaseDate < today;
  });
  for (var index = completedDeadlineQuestions.length - 1; index >= 0; index -= 1) {
    if (completedSet[completedDeadlineQuestions[index].questionId]) break;
    consecutiveMissedDays += 1;
  }

  var targetDays = releasedQuestions.length;
  var completedDays = completedIds.length;
  var correct = Object.keys(correctSet).length;
  var missedDays = Math.max(0, targetDays - completedDays);
  return {
    targetDays: targetDays,
    completedDays: completedDays,
    attempts: completedDays,
    correct: correct,
    missedDays: missedDays,
    implementationRate: targetDays ? Math.round(completedDays / targetDays * 100) : 0,
    correctRate: targetDays ? Math.round(correct / targetDays * 100) : 0,
    consecutiveMissedDays: consecutiveMissedDays,
    inactiveDays: consecutiveMissedDays,
    lastDate: allRows.length ? allRows.map(function (row) { return row.practiceDate; }).sort().pop() : null,
    history: allRows.sort(function (a, b) { return b.practiceDate.localeCompare(a.practiceDate); })
  };
}

function refreshDashboardsAction_() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { refreshed: false, busy: true };
  try {
    refreshDashboards_();
    return { refreshed: true, busy: false };
  } finally {
    lock.releaseLock();
  }
}

function refreshDashboards() {
  refreshDashboards_();
  getSpreadsheet_().toast('集計を更新しました。', '文脈勝負', 3);
}

function refreshDashboards_() {
  var students = getStudents_().filter(function (student) { return student.active; });
  var questions = getQuestions_();
  var attempts = getSubmittedRows_();
  var today = tokyoDate_();
  var activeStudentSet = {};
  students.forEach(function (student) { activeStudentSet[student.studentId] = true; });

  var studentRows = students.map(function (student) {
    return { student: student, stats: getStudentStats_(student.studentId, questions, attempts, today) };
  });
  var progressRows = studentRows.slice().sort(function (a, b) {
    return a.stats.implementationRate - b.stats.implementationRate ||
      b.stats.missedDays - a.stats.missedDays ||
      a.student.campus.localeCompare(b.student.campus, 'ja') ||
      a.student.fullName.localeCompare(b.student.fullName, 'ja');
  });
  var followRows = studentRows.filter(function (row) {
    return row.stats.consecutiveMissedDays >= 3;
  }).sort(function (a, b) {
    return b.stats.consecutiveMissedDays - a.stats.consecutiveMissedDays ||
      a.stats.implementationRate - b.stats.implementationRate ||
      a.student.campus.localeCompare(b.student.campus, 'ja') ||
      a.student.fullName.localeCompare(b.student.fullName, 'ja');
  });

  var targetDays = studentRows.length ? studentRows[0].stats.targetDays :
    questions.filter(function (question) { return question.releaseDate <= today; }).length;
  var totalTarget = students.length * targetDays;
  var totalCompleted = studentRows.reduce(function (sum, row) {
    return sum + row.stats.completedDays;
  }, 0);
  var totalCorrect = studentRows.reduce(function (sum, row) {
    return sum + row.stats.correct;
  }, 0);
  var releasedIdSet = {};
  questions.filter(function (question) {
    return question.releaseDate <= today;
  }).forEach(function (question) {
    releasedIdSet[question.questionId] = true;
  });
  var todayStudents = unique_(attempts.filter(function (row) {
    return row.practiceDate === today && activeStudentSet[row.studentId] && releasedIdSet[row.questionId];
  }).map(function (row) {
    return row.studentId;
  })).length;

  var dashboard = requiredSheet_(SHEETS.DASHBOARD);
  dashboard.getRange('B4:B8').setValues([
    [students.length],
    [targetDays],
    [todayStudents],
    [totalCompleted + ' / ' + totalTarget],
    [totalCorrect + ' / ' + totalTarget]
  ]);
  dashboard.getRange('D4:H200').clearContent();
  var questionStats = questions.map(function (question) {
    var rows = attempts.filter(function (row) {
      return row.questionId === question.questionId && activeStudentSet[row.studentId];
    });
    var correct = rows.filter(function (row) { return row.result === '正解'; }).length;
    return [
      question.questionId,
      question.title,
      rows.length,
      correct,
      rows.filter(function (row) { return row.timedOut; }).length
    ];
  });
  if (questionStats.length) dashboard.getRange(4, 4, questionStats.length, 5).setValues(questionStats);

  writeProgressSheet_(requiredSheet_(SHEETS.PROGRESS), progressRows, false);
  writeProgressSheet_(requiredSheet_(SHEETS.FOLLOW), followRows, true);
}

function writeProgressSheet_(sheet, rows, isFollow) {
  sheet.getRange('A5:I1000').clearContent();
  if (!rows.length) return;
  sheet.getRange(5, 1, rows.length, 9).setValues(rows.map(function (row) {
    var stats = row.stats;
    return [
      row.student.studentId,
      row.student.fullName,
      row.student.campus,
      stats.completedDays + ' / ' + stats.targetDays,
      stats.correct + ' / ' + stats.targetDays,
      (isFollow ? stats.consecutiveMissedDays : stats.missedDays) + '日',
      rateText_(stats.implementationRate, stats.targetDays),
      rateText_(stats.correctRate, stats.targetDays),
      stats.lastDate || '—'
    ];
  }));
  sheet.getRange('A5:A1000').setNumberFormat('@');
}

function rateText_(rate, targetDays) {
  return targetDays ? rate + '%' : '—';
}

function refreshIndividualSearch() {
  var sheet = requiredSheet_(SHEETS.INDIVIDUAL);
  var studentId = String(sheet.getRange('B3').getDisplayValue()).trim();
  sheet.getRange('B5:B12').clearContent();
  sheet.getRange('A15:F1000').clearContent();
  if (!studentId) return;

  var student = getStudent_(studentId);
  if (!student) {
    sheet.getRange('B5').setValue('登録がありません');
    return;
  }
  var stats = getStudentStats_(studentId);
  sheet.getRange('B5:B12').setValues([
    [student.fullName],
    [student.campus],
    [stats.targetDays + '日'],
    [stats.completedDays + ' / ' + stats.targetDays],
    [stats.correct + ' / ' + stats.targetDays],
    [stats.missedDays + '日'],
    [rateText_(stats.implementationRate, stats.targetDays)],
    [rateText_(stats.correctRate, stats.targetDays)]
  ]);
  if (stats.history.length) {
    sheet.getRange(15, 1, stats.history.length, 6).setValues(stats.history.map(function (row) {
      return [row.practiceDate, row.questionId, row.timedOut ? '時間切れ' : row.result, row.answer, row.correctOrder, row.elapsedSeconds];
    }));
  }
}

function validateStudentId_(rawStudentId, config) {
  var studentId = String(rawStudentId || '').trim();
  var digits = Number(config.STUDENT_ID_DIGITS);
  if (!(new RegExp('^\\d{' + digits + '}$')).test(studentId)) {
    throw appError_('生徒IDは' + digits + '桁の数字で入力してください。', 'INVALID_STUDENT_ID');
  }
  return studentId;
}

function requiredSheet_(name) {
  var sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw appError_('「' + name + '」シートが見つかりません。', 'SHEET_NOT_FOUND');
  return sheet;
}

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var active = SpreadsheetApp.getActive();
  if (!active) throw appError_('管理用スプレッドシートを確認できません。setupWorkbook を実行してください。', 'SPREADSHEET_NOT_CONFIGURED');
  return active;
}

function tokyoDate_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function formatDateValue_(value) {
  if (value instanceof Date) return Utilities.formatDate(value, 'Asia/Tokyo', 'yyyy-MM-dd');
  var text = String(value || '').trim();
  var match = text.match(/\d{4}-\d{2}-\d{2}/);
  return match ? match[0] : text;
}

function unique_(values) {
  var seen = {};
  return values.filter(function (value) {
    if (seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function appError_(message, code) {
  var error = new Error(message);
  error.code = code;
  return error;
}

function errorPayload_(error) {
  console.error(error && error.stack ? error.stack : error);
  return {
    ok: false,
    error: error && error.message ? error.message : '処理中にエラーが発生しました。',
    code: error && error.code ? error.code : 'INTERNAL_ERROR'
  };
}

function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
