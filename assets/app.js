(function () {
  'use strict';

  var root = document.getElementById('app');
  var config = window.BUNMYAKU_CONFIG || {};
  var DRAFT_KEY = 'bunmyaku-shobu-active-attempt-v4';
  var PENDING_KEY = DRAFT_KEY + '-pending';
  var labels = ['A', 'B', 'C', 'D', 'E', 'F'];
  var timerId = null;
  var submitting = false;

  var state = {
    status: null,
    studentId: '',
    verifiedStudent: null,
    attempt: null,
    answer: [],
    result: null,
    error: '',
    notice: '',
    busy: false,
    finished: false,
    lastPhase: null
  };

  function escapeHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function apiUrl() {
    var url = String(config.apiUrl || '').trim();
    if (!url || url.indexOf('__APPS_SCRIPT_') >= 0) return '';
    return url;
  }

  async function api(action, payload) {
    var url = apiUrl();
    if (!url) throw new Error('Googleスプレッドシートとの接続設定がまだ完了していません。');
    var controller = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timeoutId = controller ? window.setTimeout(function () { controller.abort(); }, 15000) : null;
    var response;
    try {
      if (action === 'status') {
        var getOptions = {
          method: 'GET',
          cache: 'no-store',
          redirect: 'follow'
        };
        if (controller) getOptions.signal = controller.signal;
        response = await fetch(url + '?action=status&t=' + Date.now(), getOptions);
      } else {
        var postOptions = {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(Object.assign({ action: action }, payload || {})),
          redirect: 'follow'
        };
        if (controller) postOptions.signal = controller.signal;
        response = await fetch(url, postOptions);
      }
      if (!response.ok) throw new Error('通信に失敗しました。しばらくしてからもう一度お試しください。');
      var body = await response.json();
      if (!body.ok) {
        var error = new Error(body.error || '処理に失敗しました。');
        error.code = body.code;
        throw error;
      }
      return body.data;
    } catch (error) {
      if (error && error.name === 'AbortError') {
        throw new Error('通信に時間がかかっています。自動的に再接続します。');
      }
      throw error;
    } finally {
      if (timeoutId) window.clearTimeout(timeoutId);
    }
  }

function periodText() {
    var period = state.status && state.status.period;
    if (!period) return '2026.10.01 — 10.31';
    return period.start.replace(/-/g, '.') + ' — ' + period.end.slice(5).replace('-', '.');
  }

  function timing() {
    if (state.attempt && state.attempt.testMode) {
      var base = state.attempt.timing || { answerSeconds: 60 };
      return { readingSeconds: 0, answerSeconds: Number(base.answerSeconds || 60), totalSeconds: Number(base.answerSeconds || 60) };
    }
    if (state.attempt && state.attempt.timing) return state.attempt.timing;
    if (state.status && state.status.timing) return state.status.timing;
    return { readingSeconds: 180, answerSeconds: 60, totalSeconds: 240 };
  }

  function phaseInfo() {
    if (!state.attempt) return { phase: 'none', remaining: 0, elapsed: 0 };
    var elapsed = Math.max(0, (Date.now() - new Date(state.attempt.startedAt).getTime()) / 1000);
    var t = timing();
    var phase = elapsed < t.readingSeconds ? 'reading' : elapsed < t.totalSeconds ? 'answer' : 'locked';
    var end = phase === 'reading' ? t.readingSeconds : t.totalSeconds;
    return { phase: phase, remaining: Math.max(0, Math.ceil(end - elapsed)), elapsed: elapsed };
  }

  function formatClock(seconds) {
    return Math.floor(seconds / 60) + ':' + String(seconds % 60).padStart(2, '0');
  }

  function persistDraft() {
    if (!state.attempt) return;
    localStorage.setItem(DRAFT_KEY, JSON.stringify({
      attempt: state.attempt,
      answer: state.answer,
      studentId: state.studentId
    }));
  }

  function clearDraft() {
    localStorage.removeItem(DRAFT_KEY);
    localStorage.removeItem(PENDING_KEY);
  }

  function renderEntry() {
    stopTimer();
    var digits = state.status ? state.status.studentIdDigits : 8;
    var student = state.verifiedStudent;
    root.innerHTML = [
      '<main class="entry-page"><section class="entry-shell">',
      '<div class="brand-lockup">',
      '<span class="brand-kicker">6段落並べ替えトレーニング</span>',
      '<h1>文脈<span>勝負</span></h1>',
      '<p>つながりを読み、順序を見抜く。<br>4分間の読解トレーニング。</p>',
      '</div>',
      '<div class="entry-card">',
      '<div class="today-mark"><span>実施期間</span><strong>' + escapeHtml(periodText()) + '</strong></div>',
      student ? studentConfirmationHtml(student) : [
        '<form id="student-form">',
        '<label for="student-id">生徒ID</label>',
        '<input id="student-id" inputmode="numeric" pattern="[0-9]*" maxlength="' + digits + '" value="' + escapeHtml(state.studentId) + '" placeholder="' + digits + '桁の数字" autocomplete="off">',
        '<button class="primary-button" type="submit" ' + (state.busy || state.studentId.length !== digits ? 'disabled' : '') + '>',
        state.busy ? '確認中…' : '生徒IDを確認する',
        '</button></form>'
      ].join(''),
      state.notice ? '<p class="notice">' + escapeHtml(state.notice) + '</p>' : '',
      state.error ? '<p class="error" role="alert">' + escapeHtml(state.error) + '</p>' : '',
      '<div class="rules-strip"><span><b>3分</b>読解</span><i></i><span><b>60秒</b>並べ替え</span><i></i><span><b>1日</b>1問</span></div>',
      '</div></section></main>'
    ].join('');

    var input = document.getElementById('student-id');
    if (input) {
      input.addEventListener('input', function (event) {
        state.studentId = event.target.value.replace(/\D/g, '').slice(0, digits);
        state.error = '';
        var button = document.querySelector('#student-form button');
        if (button) button.disabled = state.studentId.length !== digits;
      });
      document.getElementById('student-form').addEventListener('submit', verifyStudent);
    }
    var start = document.getElementById('confirm-start');
    if (start) start.addEventListener('click', beginAttempt);
    var back = document.getElementById('confirm-back');
    if (back) back.addEventListener('click', function () {
      state.verifiedStudent = null;
      state.error = '';
      renderEntry();
    });
  }

  function studentConfirmationHtml(student) {
    return [
      '<div class="student-confirm"><span>生徒情報を確認してください</span>',
      '<strong>' + escapeHtml(student.fullName) + '</strong>',
      student.campus ? '<p>' + escapeHtml(student.campus) + '</p>' : '',
      '<div class="student-streak"><span>現在の連続実施</span><strong>' + Number(student.streakDays || 0) + '<small>日</small></strong></div>',
      student.testMode ? '<p class="test-note">テスト実施｜提出内容は実施記録に保存されます</p>' : '',
      '<div class="confirm-actions">',
      '<button id="confirm-start" type="button" class="confirm-start" ' + (state.busy ? 'disabled' : '') + '>' + (state.busy ? '開始準備中…' : 'この生徒で始める') + '</button>',
      '<button id="confirm-back" type="button" class="confirm-back" ' + (state.busy ? 'disabled' : '') + '>入力し直す</button>',
      '</div></div>'
    ].join('');
  }

  async function verifyStudent(event) {
    event.preventDefault();
    state.busy = true;
    state.error = '';
    state.notice = '';
    renderEntry();
    try {
      state.verifiedStudent = await api('verifyStudent', { studentId: state.studentId });
    } catch (error) {
      state.error = error.message;
    } finally {
      state.busy = false;
      renderEntry();
    }
  }

  async function beginAttempt() {
    state.busy = true;
    state.error = '';
    renderEntry();
    try {
      state.attempt = await api('startAttempt', { studentId: state.studentId });
      state.answer = [];
      state.result = null;
      state.finished = false;
      state.lastPhase = null;
      state.busy = false;
      persistDraft();
      renderChallenge();
      startTimer();
    } catch (error) {
      state.error = error.message;
      state.busy = false;
      renderEntry();
    }
  }

  function renderChallenge() {
    if (!state.attempt) return renderEntry();
    var info = phaseInfo();
    state.lastPhase = info.phase;
    var question = state.attempt.question;
    root.innerHTML = [
      '<main class="challenge-page phase-' + info.phase + '">',
      '<header class="challenge-header"><div><span>文脈勝負</span><strong>' + escapeHtml(question.title) + '</strong>',
      state.attempt.testMode ? '<small>テスト実施｜記録確認用（提出内容は保存されます）</small>' : '',
      '</div><div id="phase-pill" class="phase-pill ' + info.phase + '"><span id="phase-label">' + (submitting ? '送信中' : info.phase === 'reading' ? '読解中' : '解答中') + '</span><strong id="timer">' + (submitting ? '—' : formatClock(info.remaining)) + '</strong></div></header>',
      '<div class="phase-banner ' + info.phase + '" role="status">',
      submitting
        ? '<strong>送信中</strong><span>解答を受け付けています。画面を閉じずにお待ちください。</span>'
        : info.phase === 'reading'
          ? '<strong>読解中</strong><span>今は文章のつながりを考える時間です。解答操作はまだできません。</span>'
          : '<strong>解答中</strong><span>' + (state.attempt.testMode ? 'テスト用IDのため、すぐに並べ替えできます。' : '画面が黄色に変わりました。A～Fを正しい順に選んでください。') + '</span>',
      '</div>',
      '<section class="paragraph-board" aria-label="問題文">',
      question.paragraphs.map(function (paragraph) {
        return '<article class="paragraph-card"><span>' + escapeHtml(paragraph.label) + '</span><p>' + escapeHtml(paragraph.text) + '</p></article>';
      }).join(''),
      '</section>',
      '<section id="answer-dock" class="answer-dock ' + (info.phase === 'reading' ? 'is-locked' : '') + '">',
      answerDockHtml(info.phase),
      state.error ? '<p class="dock-error" role="alert">' + escapeHtml(state.error) + '</p>' : '',
      '</section></main>'
    ].join('');
    bindAnswerControls();
  }

  function answerDockHtml(phase) {
    var reading = phase === 'reading';
    var editable = phase === 'answer' && !submitting;
    var canSubmit = !submitting && (
      (phase === 'answer' && state.answer.length === 6) ||
      Boolean(state.error)
    );
    return [
      '<div class="answer-heading"><strong>' + (submitting ? '解答を送信しています' : reading ? '3分間で文章の流れを考えよう' : '正しい順に記号を選ぼう') + '</strong>',
      '<span>' + (submitting ? '送信が完了すると結果画面へ切り替わります' : reading ? '180秒後に解答できます' : '入力した枠を押すと取り消せます') + '</span></div>',
      '<div class="answer-controls"><div class="answer-slots" aria-label="解答枠">',
      Array.from({ length: 6 }, function (_, index) {
        var value = state.answer[index] || '';
        return '<button type="button" data-slot="' + index + '" ' + (!editable || index >= state.answer.length ? 'disabled' : '') + ' aria-label="' + (index + 1) + '番目、' + (value || '未入力') + '"><small>' + (index + 1) + '</small><b>' + escapeHtml(value) + '</b></button>';
      }).join(''),
      '</div><div class="label-buttons">',
      labels.map(function (label) {
        return '<button type="button" data-label="' + label + '" ' + (!editable || state.answer.indexOf(label) >= 0 ? 'disabled' : '') + '>' + label + '</button>';
      }).join(''),
      '</div><button id="submit-answer" class="submit-answer" type="button" ' + (!canSubmit ? 'disabled' : '') + '>' + (submitting ? '提出中…' : state.error ? '提出を再送する' : '解答する') + '</button></div>'
    ].join('');
  }

function bindAnswerControls() {
    document.querySelectorAll('[data-label]').forEach(function (button) {
      button.addEventListener('click', function () {
        if (submitting || phaseInfo().phase !== 'answer' || state.answer.indexOf(button.dataset.label) >= 0) return;
        state.answer.push(button.dataset.label);
        persistDraft();
        renderChallenge();
      });
    });
    document.querySelectorAll('[data-slot]').forEach(function (button) {
      button.addEventListener('click', function () {
        if (submitting || phaseInfo().phase !== 'answer') return;
        state.answer.splice(Number(button.dataset.slot), 1);
        persistDraft();
        renderChallenge();
      });
    });
    var submit = document.getElementById('submit-answer');
    if (submit) submit.addEventListener('click', function () { submitAttempt(new Date().toISOString()); });
  }

  function startTimer() {
    stopTimer();
    timerId = window.setInterval(updateTimer, 250);
    updateTimer();
  }

  function stopTimer() {
    if (timerId) window.clearInterval(timerId);
    timerId = null;
  }

  function updateTimer() {
    if (!state.attempt || state.result) return stopTimer();
    var info = phaseInfo();
    if (info.phase !== state.lastPhase) {
      renderChallenge();
      if (info.phase === 'locked') submitAttempt(new Date().toISOString());
      return;
    }
    var timer = document.getElementById('timer');
    if (timer) timer.textContent = formatClock(info.remaining);
    if (info.phase === 'locked') submitAttempt(new Date().toISOString());
  }

  async function submitAttempt(lockedAt) {
    if (submitting || !state.attempt) return;
    submitting = true;
    state.busy = true;
    state.error = '';
    stopTimer();
    var payload = {
      attemptId: state.attempt.attemptId,
      attemptToken: state.attempt.attemptToken,
      answer: state.answer.slice(),
      lockedAt: lockedAt
    };
    localStorage.setItem(PENDING_KEY, JSON.stringify(payload));
    renderChallenge();
    var lastError = null;
    for (var attemptNumber = 1; attemptNumber <= 5; attemptNumber += 1) {
      try {
        var response = await api('submitAttempt', payload);
        state.result = response.result;
        state.busy = false;
        submitting = false;
        clearDraft();
        stopTimer();
        renderResult();
        window.setTimeout(function () {
          api('refreshDashboards', {}).catch(function () {});
        }, 0);
        return;
      } catch (error) {
        lastError = error;
        if (attemptNumber < 5) await new Promise(function (resolve) { window.setTimeout(resolve, 3000); });
      }
    }
    state.busy = false;
    submitting = false;
    state.error = (lastError ? lastError.message : '提出できませんでした。') + ' 通信を確認して「解答する」をもう一度押してください。';
    renderChallenge();
  }

  function highlightedText(text, highlights) {
    var highlight = (highlights || []).filter(Boolean).sort(function (a, b) { return b.length - a.length; }).filter(function (item) {
      return text.indexOf(item) >= 0;
    })[0];
    if (!highlight) return escapeHtml(text);
    var index = text.indexOf(highlight);
    return escapeHtml(text.slice(0, index)) + '<mark class="review-underline">' + escapeHtml(highlight) + '</mark>' + escapeHtml(text.slice(index + highlight.length));
  }

  function renderResult() {
    if (!state.result || !state.attempt) return;
    var result = state.result;
    var question = state.attempt.question;
    var title = result.timedOut ? '時間切れ' : result.correct ? '正解！' : 'もう一歩';
    var lead = result.timedOut ? '制限時間になりました。正しい流れを確認しましょう。' : result.correct ? '文章の流れを正確につかめました。' : '一番大切なつながりだけ、確認しましょう。';
    root.innerHTML = [
      '<main class="result-page"><section class="result-card ' + (result.correct ? 'is-correct' : 'is-wrong') + '">',
      '<p class="eyebrow">文脈勝負｜' + escapeHtml(question.releaseDate.replace(/-/g, '.')) + '</p>',
      '<h1>' + title + '</h1><p class="result-lead">' + lead + '</p>',
      '<div class="result-streak"><span>今日で</span><strong>' + Number(result.streakDays || 0) + '</strong><b>日連続！</b></div>',
      '<div class="answer-compare"><div><span>あなたの答え</span><strong>' + escapeHtml(result.answer.length ? result.answer.join(' → ') : '未完成') + '</strong></div>',
      '<div><span>正しい順番</span><strong>' + escapeHtml(result.correctOrder.join(' → ')) + '</strong></div></div>',
      !result.correct && result.feedbackRule ? '<div class="feedback-box"><span>ここを確認しよう</span><p>' + escapeHtml(result.feedbackRule.explanation) + '</p></div>' : '',
      !result.correct && result.review ? reviewHtml(result.review, question.paragraphs) : '',
      '<p class="completion-note">' + (state.attempt.testMode ? 'テスト実施の内容を実施記録に保存しました。' : '本日の文脈勝負は終了です。再挑戦はできません。') + '</p>',
      '<button id="finish-learning" class="finish-learning" type="button">学習を終える</button>',
      '</section></main>'
    ].join('');
    document.getElementById('finish-learning').addEventListener('click', function () {
      state.finished = true;
      renderFinished();
    });
  }

  function reviewHtml(review, paragraphs) {
    return [
      '<section class="review-section"><div class="review-heading"><div><span>本文に戻って確認</span><h2>下線部のつながりを見よう</h2></div><p>今回のポイント</p></div>',
      '<div class="review-paragraphs">',
      paragraphs.map(function (paragraph) {
        var focused = review.focusLabels.indexOf(paragraph.label) >= 0;
        return '<article class="review-paragraph ' + (focused ? 'is-focus' : '') + '"><b>' + escapeHtml(paragraph.label) + '</b><p>' + highlightedText(paragraph.text, review.highlights[paragraph.label]) + '</p></article>';
      }).join(''),
      '</div><div class="advice-box"><strong>次に生かすアドバイス</strong><p>' + escapeHtml(review.advice) + '</p></div></section>'
    ].join('');
  }

  function renderFinished() {
    root.innerHTML = [
      '<main class="finished-page"><section class="finished-card">',
      '<div class="finished-mark" aria-hidden="true">✓</div>',
      '<p class="eyebrow">文脈勝負｜本日のトレーニング</p>',
      '<h1>今日の学習は終了しました</h1>',
      state.result ? '<div class="finished-streak"><span>連続実施</span><strong>' + Number(state.result.streakDays || 0) + '<small>日</small></strong></div>' : '',
      '<p>文章のつながりを意識して読んだ経験を、次の問題にも生かしましょう。</p>',
      '<div class="finished-actions">',
      '<button id="to-entry" class="finish-primary" type="button">ID入力画面へ戻る</button>',
      '</div></section></main>'
    ].join('');
    document.getElementById('to-entry').addEventListener('click', function () { reset(false); });
  }

  function reset(keepTestId) {
    var previousTest = state.attempt && state.attempt.testMode;
    clearDraft();
    state.attempt = null;
    state.answer = [];
    state.result = null;
    state.finished = false;
    state.error = '';
    state.notice = '';
    state.busy = false;
    state.verifiedStudent = null;
    if (!keepTestId || !previousTest) state.studentId = '';
    if (keepTestId && previousTest) {
      state.verifiedStudent = { studentId: state.studentId, fullName: 'テスト生徒', campus: '', testMode: true, streakDays: 0 };
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
    renderEntry();
  }

  function renderSetup(message) {
    root.innerHTML = '<main class="setup-page"><section class="setup-card"><h1>文脈勝負</h1><p>' + escapeHtml(message) + '</p><code>管理者の初期設定後に利用できます。</code></section></main>';
  }

  async function bootstrap() {
    if (!apiUrl()) {
      renderSetup('生徒用画面はGitHubへ配置済みです。現在、Googleスプレッドシートとの接続設定待ちです。');
      return;
    }
    try {
      state.status = await api('status');
      if (!state.status.available) state.notice = '次回の実施期間は ' + state.status.period.start + '～' + state.status.period.end + ' です。';
    } catch (error) {
      renderSetup('管理用スプレッドシートへ接続できません。管理者がApps Scriptの公開設定を確認してください。');
      return;
    }

    var saved = localStorage.getItem(DRAFT_KEY);
    if (saved) {
      try {
        var draft = JSON.parse(saved);
        state.attempt = draft.attempt;
        state.answer = Array.isArray(draft.answer) ? draft.answer : [];
        state.studentId = draft.studentId || '';
        renderChallenge();
        startTimer();
        var pending = localStorage.getItem(PENDING_KEY);
        if (pending) {
          var payload = JSON.parse(pending);
          submitAttempt(payload.lockedAt || new Date().toISOString());
        }
        return;
      } catch (error) {
        clearDraft();
      }
    }
    renderEntry();
  }

  bootstrap();
})();
