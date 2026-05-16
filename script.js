(() => {
  "use strict";

  const STORAGE_KEY = "infinite-tic-tac-toe:v1";
  const CELLS = Array.from({ length: 9 }, (_, i) => i);
  const LINES = [
    [0, 1, 2], [3, 4, 5], [6, 7, 8],
    [0, 3, 6], [1, 4, 7], [2, 5, 8],
    [0, 4, 8], [2, 4, 6]
  ];
  const LINE_META = {
    "0,1,2": { top: "16.66%", left: "8%", rotate: 0, width: "84%" },
    "3,4,5": { top: "50%", left: "8%", rotate: 0, width: "84%" },
    "6,7,8": { top: "83.33%", left: "8%", rotate: 0, width: "84%" },
    "0,3,6": { top: "50%", left: "-25%", rotate: 90, width: "84%" },
    "1,4,7": { top: "50%", left: "8%", rotate: 90, width: "84%" },
    "2,5,8": { top: "50%", left: "41%", rotate: 90, width: "84%" },
    "0,4,8": { top: "50%", left: "3%", rotate: 45, width: "94%" },
    "2,4,6": { top: "50%", left: "3%", rotate: -45, width: "94%" }
  };
  const TRANSFORMS = [
    [0,1,2,3,4,5,6,7,8], [6,3,0,7,4,1,8,5,2], [8,7,6,5,4,3,2,1,0], [2,5,8,1,4,7,0,3,6],
    [2,1,0,5,4,3,8,7,6], [6,7,8,3,4,5,0,1,2], [0,3,6,1,4,7,2,5,8], [8,5,2,7,4,1,6,3,0]
  ];
  const COLLECTION_TOTAL = 8;
  const ACHIEVEMENTS = [
    { id: "firstWin", icon: "🔥", title: "初勝利", desc: "プレイヤーが1勝する。" },
    { id: "streak3", icon: "⚡", title: "3連勝", desc: "3連勝を達成する。" },
    { id: "comeback", icon: "🛡", title: "逆転勝利", desc: "相手リーチを防いだ次ターン以内に勝利。" },
    { id: "trap", icon: "🕸", title: "トラップ勝利", desc: "フォーク（ダブルリーチ）を作って勝利。" },
    { id: "vanish", icon: "🌌", title: "消滅逆転", desc: "相手駒の消滅で防御不能にして勝利。" },
    { id: "complete", icon: "👑", title: "コンプリート", desc: "1〜5をすべて解除。" }
  ];

  const defaultSave = () => ({
    stats: { playerScore: 0, cpuScore: 0, totalGames: 0, wins: 0, streak: 0, bestStreak: 0 },
    achievements: {}, collection: {}, seenNew: {},
    settings: { bgmVolume: 0.16, seVolume: 0.7, difficulty: "NORMAL", effects: "normal", colorAssist: false, shake: true, muted: false }
  });

  const state = {
    screen: "title", phase: "playing", turn: "player", turnNumber: 1,
    board: Array(9).fill(null), queues: { player: [], cpu: [] }, winner: null, winningLine: null,
    save: loadSave(), cpuTimer: null, roundFlags: { blockedAtTurn: -99, hadFork: false, vanishOpportunity: false },
    audio: null, lastHint: "中央と角は消滅後も形を作り直しやすい要所です。"
  };

  const $ = (id) => document.getElementById(id);
  const dom = {};

  document.addEventListener("DOMContentLoaded", init);

  function init() {
    cacheDom();
    buildBoard();
    bindEvents();
    applySettingsClasses();
    startAmbientTone();
    newRound(false);
    renderAll();
    goTo("title");
  }

  function cacheDom() {
    ["titleScreen","gameScreen","endScreen","achievementsScreen","collectionScreen","settingsScreen","board","boardFrame","turnChip","turnNumber","winRate","playerScore","cpuScore","streak","totalGames","achievementProgressText","achievementProgressBar","collectionProgressText","collectionProgressBar","miniAchievements","achievementGrid","collectionGrid","achievementSummary","collectionSummary","microFeed","thinkingIndicator","winLine","resultCard","resultTitle","resultText","resultMetrics","heroWinRate","heroGames","heroCollection","nextRewardHint","threatPanel","settingsSnap","screenFlash","toastStack","bgmVolume","seVolume","difficultySelect","effectsSelect","colorAssist","shakeEnabled","soundToggle"].forEach(id => dom[id] = $(id));
  }

  function bindEvents() {
    document.querySelectorAll("[data-screen]").forEach(btn => btn.addEventListener("click", () => goTo(btn.dataset.screen)));
    $("startButton").addEventListener("click", () => goTo("game"));
    $("continueButton").addEventListener("click", () => { newRound(true); goTo("game"); });
    $("newRoundButton").addEventListener("click", () => newRound(true));
    $("hintButton").addEventListener("click", showHint);
    $("resetProgressButton").addEventListener("click", () => {
      if (!confirm("全ての進行データを初期化しますか？")) return;
      localStorage.removeItem(STORAGE_KEY); location.reload();
    });
    dom.soundToggle.addEventListener("click", () => updateSetting("muted", !state.save.settings.muted));
    [[dom.bgmVolume,"bgmVolume"],[dom.seVolume,"seVolume"],[dom.difficultySelect,"difficulty"],[dom.effectsSelect,"effects"],[dom.colorAssist,"colorAssist"],[dom.shakeEnabled,"shake"]].forEach(([el,key]) => {
      el.addEventListener("input", () => updateSetting(key, el.type === "checkbox" ? el.checked : el.value));
    });
    window.addEventListener("pointerdown", unlockAudio, { once: true });
  }

  function buildBoard() {
    dom.board.innerHTML = "";
    CELLS.forEach(index => {
      const cell = document.createElement("button");
      cell.className = "cell";
      cell.type = "button";
      cell.dataset.index = index;
      cell.setAttribute("aria-label", `${index + 1}マス目`);
      cell.addEventListener("click", () => handlePlayerMove(index));
      dom.board.appendChild(cell);
    });
  }

  function loadSave() {
    try { return { ...defaultSave(), ...JSON.parse(localStorage.getItem(STORAGE_KEY) || "{}") }; }
    catch { return defaultSave(); }
  }
  function persist() { localStorage.setItem(STORAGE_KEY, JSON.stringify(state.save)); }

  function updateSetting(key, value) {
    if (key === "bgmVolume" || key === "seVolume") value = Number(value);
    state.save.settings[key] = value;
    persist(); applySettingsClasses(); renderAll(); playSound("ui");
  }

  function applySettingsClasses() {
    document.body.classList.toggle("color-assist", state.save.settings.colorAssist);
    document.body.classList.toggle("low-effects", state.save.settings.effects === "low");
    document.body.classList.toggle("high-effects", state.save.settings.effects === "high");
    if (dom.soundToggle) dom.soundToggle.classList.toggle("off", state.save.settings.muted);
    if (state.audio?.bgmGain) state.audio.bgmGain.gain.value = state.save.settings.muted ? 0 : state.save.settings.bgmVolume;
  }

  function goTo(screen) {
    state.screen = screen;
    document.querySelectorAll(".screen").forEach(s => s.classList.remove("is-visible"));
    $(`${screen}Screen`)?.classList.add("is-visible");
    document.querySelectorAll(".nav-tabs button").forEach(b => b.classList.toggle("active", b.dataset.screen === screen));
    renderAll();
  }

  function newRound(withSound) {
    clearTimeout(state.cpuTimer);
    Object.assign(state, { phase: "playing", turn: "player", turnNumber: 1, board: Array(9).fill(null), queues: { player: [], cpu: [] }, winner: null, winningLine: null, roundFlags: { blockedAtTurn: -99, hadFork: false, vanishOpportunity: false } });
    dom.winLine.className = "win-line";
    dom.boardFrame.classList.remove("victory", "defeat");
    if (withSound) { playSound("ui"); pushFeed("新しい時間軸を開始。最古駒の消滅を読み切れ。"); }
    renderAll();
  }

  function handlePlayerMove(index) {
    if (state.phase !== "playing" || state.turn !== "player" || state.board[index]) return;
    const beforeCpuThreats = immediateWinningMoves(snapshot(), "cpu").length;
    const move = applyMove(state, "player", index, true);
    if (!move) return;
    playSound("placePlayer");
    spawnPlacementEffect(index, "player");
    if (beforeCpuThreats && lineWasBlocked(index, "cpu")) state.roundFlags.blockedAtTurn = state.turnNumber;
    state.roundFlags.hadFork = state.roundFlags.hadFork || countForks(snapshot(), "player") >= 2;
    postMove("player", move.removed);
  }

  function postMove(actor, removed) {
    if (removed) { spawnVanishEffect(removed, actor); playSound("vanish"); }
    const result = getWinner(state.board);
    if (result) return finishRound(result.player, result.line);
    state.turn = actor === "player" ? "cpu" : "player";
    state.turnNumber += actor === "cpu" ? 1 : 0;
    renderAll();
    if (state.turn === "cpu") scheduleCpuMove();
  }

  function scheduleCpuMove() {
    dom.thinkingIndicator.classList.add("show");
    const delay = state.save.settings.difficulty === "HARD" ? 520 : 680;
    state.cpuTimer = setTimeout(() => {
      dom.thinkingIndicator.classList.remove("show");
      if (state.phase !== "playing" || state.turn !== "cpu") return;
      const index = chooseCpuMove();
      const move = applyMove(state, "cpu", index, true);
      playSound("placeCpu");
      spawnPlacementEffect(index, "cpu");
      postMove("cpu", move?.removed);
    }, delay);
  }

  function snapshot() { return { board: [...state.board], queues: { player: [...state.queues.player], cpu: [...state.queues.cpu] } }; }

  function applyMove(target, player, index, mutate) {
    if (target.board[index]) return null;
    const nextBoard = mutate ? target.board : [...target.board];
    const nextQueues = mutate ? target.queues : { player: [...target.queues.player], cpu: [...target.queues.cpu] };
    let removed = null;
    if (nextQueues[player].length >= 3) {
      removed = nextQueues[player].shift();
      nextBoard[removed] = null;
    }
    nextBoard[index] = player;
    nextQueues[player].push(index);
    return { board: nextBoard, queues: nextQueues, removed };
  }

  function simulate(pos, player, index) {
    const copy = { board: [...pos.board], queues: { player: [...pos.queues.player], cpu: [...pos.queues.cpu] } };
    applyMove(copy, player, index, true);
    return copy;
  }

  function chooseCpuMove() {
    const level = state.save.settings.difficulty;
    const legal = CELLS.filter(i => !state.board[i]);
    if (level === "EASY" && Math.random() < .48) return randomItem(legal);
    const tactical = findTacticalMove(snapshot(), "cpu", level);
    if (tactical != null && !(level === "EASY" && Math.random() < .22)) return tactical;
    if (level === "HARD") return minimaxMove(snapshot(), 5);
    if (level === "NORMAL") return weightedMove(snapshot(), "cpu");
    return randomItem(legal);
  }

  function findTacticalMove(pos, me, level) {
    const enemy = other(me);
    const legal = CELLS.filter(i => !pos.board[i]);
    const win = legal.find(i => getWinner(simulate(pos, me, i).board)?.player === me);
    if (win != null) return win;
    const block = legal.find(i => getWinner(simulate(pos, enemy, i).board)?.player === enemy);
    if (block != null) return block;
    if (level !== "EASY") {
      const fork = legal.find(i => countForks(simulate(pos, me, i), me) >= 2);
      if (fork != null) return fork;
      const stopFork = legal.find(i => countForks(simulate(pos, enemy, i), enemy) < 2 && countForks(simulate(pos, me, i), enemy) === 0);
      if (stopFork != null) return stopFork;
    }
    if (level === "HARD") {
      const vanishExploit = legal.find(i => createsVanishTrap(simulate(pos, me, i), me));
      if (vanishExploit != null) return vanishExploit;
    }
    return null;
  }

  function minimaxMove(pos, depth) {
    let best = -Infinity, moves = [];
    CELLS.filter(i => !pos.board[i]).forEach(i => {
      const score = minimax(simulate(pos, "cpu", i), depth - 1, false, -Infinity, Infinity);
      if (score > best) { best = score; moves = [i]; }
      else if (score === best) moves.push(i);
    });
    return tieBreak(moves, pos, "cpu");
  }

  function minimax(pos, depth, maximizing, alpha, beta) {
    const winner = getWinner(pos.board);
    if (winner) return winner.player === "cpu" ? 1000 + depth : -1000 - depth;
    if (depth <= 0) return evaluatePosition(pos, "cpu");
    const player = maximizing ? "cpu" : "player";
    const legal = CELLS.filter(i => !pos.board[i]);
    if (maximizing) {
      let value = -Infinity;
      for (const i of legal) { value = Math.max(value, minimax(simulate(pos, player, i), depth - 1, false, alpha, beta)); alpha = Math.max(alpha, value); if (beta <= alpha) break; }
      return value;
    }
    let value = Infinity;
    for (const i of legal) { value = Math.min(value, minimax(simulate(pos, player, i), depth - 1, true, alpha, beta)); beta = Math.min(beta, value); if (beta <= alpha) break; }
    return value;
  }

  function weightedMove(pos, me) {
    const legal = CELLS.filter(i => !pos.board[i]);
    let best = -Infinity, moves = [];
    legal.forEach(i => {
      const score = evaluatePosition(simulate(pos, me, i), me) + [4,0,2,6,8,1,3,5,7].indexOf(i) * -0.1;
      if (score > best) { best = score; moves = [i]; }
      else if (Math.abs(score - best) < .001) moves.push(i);
    });
    return tieBreak(moves, pos, me);
  }

  function evaluatePosition(pos, me) {
    const enemy = other(me);
    let score = 0;
    for (const line of LINES) {
      const mine = line.filter(i => pos.board[i] === me).length;
      const theirs = line.filter(i => pos.board[i] === enemy).length;
      if (mine && !theirs) score += mine === 2 ? 28 : 6;
      if (theirs && !mine) score -= theirs === 2 ? 34 : 7;
    }
    score += countForks(pos, me) * 18 - countForks(pos, enemy) * 24;
    if (pos.board[4] === me) score += 10;
    if (pos.board[4] === enemy) score -= 10;
    [0,2,6,8].forEach(c => { if (pos.board[c] === me) score += 4; if (pos.board[c] === enemy) score -= 4; });
    score += vanishTimingScore(pos, me) - vanishTimingScore(pos, enemy);
    return score;
  }

  function vanishTimingScore(pos, me) {
    const enemy = other(me);
    let score = 0;
    const enemyOldest = pos.queues[enemy][0];
    if (pos.queues[enemy].length === 3 && enemyOldest != null) {
      for (const line of LINES) {
        if (!line.includes(enemyOldest)) continue;
        const mine = line.filter(i => pos.board[i] === me).length;
        const theirs = line.filter(i => pos.board[i] === enemy).length;
        const empty = line.filter(i => !pos.board[i]).length;
        if (mine === 2 && theirs === 1 && empty === 0) score += 26;
      }
    }
    return score;
  }

  function createsVanishTrap(pos, me) { return vanishTimingScore(pos, me) >= 26 || immediateWinningMoves(pos, me).length >= 2; }
  function tieBreak(moves, pos, me) { return [4,0,2,6,8,1,3,5,7].find(i => moves.includes(i)) ?? randomItem(CELLS.filter(i => !pos.board[i])); }
  function randomItem(arr) { return arr[Math.floor(Math.random() * arr.length)]; }
  function other(player) { return player === "player" ? "cpu" : "player"; }

  function getWinner(board) {
    for (const line of LINES) {
      const [a,b,c] = line;
      if (board[a] && board[a] === board[b] && board[b] === board[c]) return { player: board[a], line };
    }
    return null;
  }

  function immediateWinningMoves(pos, player) {
    return CELLS.filter(i => !pos.board[i] && getWinner(simulate(pos, player, i).board)?.player === player);
  }

  function countForks(pos, player) { return immediateWinningMoves(pos, player).length; }

  function lineWasBlocked(index, attackingPlayer) {
    return LINES.some(line => line.includes(index) && line.filter(i => state.board[i] === attackingPlayer).length === 2 && line.filter(i => state.board[i] === "player").length === 1);
  }

  function finishRound(winner, line) {
    state.phase = "ended"; state.winner = winner; state.winningLine = line;
    const playerWon = winner === "player";
    state.save.stats.totalGames += 1;
    state.save.stats[playerWon ? "playerScore" : "cpuScore"] += 1;
    if (playerWon) { state.save.stats.wins += 1; state.save.stats.streak += 1; state.save.stats.bestStreak = Math.max(state.save.stats.bestStreak, state.save.stats.streak); }
    else state.save.stats.streak = 0;
    if (playerWon) recordCollection();
    const unlocked = evaluateAchievements(playerWon);
    persist(); renderAll(); showWinLine(line, winner);
    dom.boardFrame.classList.add(playerWon ? "victory" : "defeat");
    flash(playerWon ? "win" : "lose");
    playSound(playerWon ? "win" : "lose");
    if (!playerWon && state.save.settings.shake && navigator.vibrate) navigator.vibrate([35, 35, 50]);
    unlocked.forEach(a => showToast(`${a.icon} ${a.title} 解除`, a.desc, true));
    setTimeout(() => showResult(playerWon, unlocked), 900);
  }

  function evaluateAchievements(playerWon) {
    const unlocked = [];
    const unlock = (id) => {
      if (!state.save.achievements[id]) {
        state.save.achievements[id] = Date.now();
        const ach = ACHIEVEMENTS.find(a => a.id === id);
        unlocked.push(ach); state.save.seenNew[id] = true; playSound("achievement");
      }
    };
    if (playerWon) {
      unlock("firstWin");
      if (state.save.stats.streak >= 3) unlock("streak3");
      if (state.turnNumber - state.roundFlags.blockedAtTurn <= 1) unlock("comeback");
      if (state.roundFlags.hadFork) unlock("trap");
      if (state.roundFlags.vanishOpportunity || vanishTimingScore(snapshot(), "player") >= 26) unlock("vanish");
    }
    if (ACHIEVEMENTS.slice(0,5).every(a => state.save.achievements[a.id])) unlock("complete");
    return unlocked;
  }

  function recordCollection() {
    const key = canonicalBoardKey(state.board);
    if (!state.save.collection[key]) {
      state.save.collection[key] = { board: state.board.map(v => v ? v[0] : "-"), at: Date.now(), line: state.winningLine };
      state.save.seenNew[`col:${key}`] = true;
      showToast("NEW ARCHIVE", "新しい勝利盤面を図鑑に登録しました。", false);
    }
  }

  function canonicalBoardKey(board) {
    const symbols = board.map(v => v === "player" ? "P" : v === "cpu" ? "C" : "-");
    return TRANSFORMS.map(map => map.map(i => symbols[i]).join("")).sort()[0];
  }

  function showResult(playerWon, unlocked) {
    dom.resultCard.classList.toggle("defeat", !playerWon);
    dom.resultTitle.textContent = playerWon ? "VICTORY" : "REVENGE?";
    dom.resultText.textContent = playerWon ? "時間軸を制圧。次の未発見パターンまであと少しです。" : "CPUが消滅タイミングを利用しました。次は最古駒を逆手に取りましょう。";
    dom.resultMetrics.innerHTML = `
      <span><b>${state.save.stats.streak}</b>連勝</span>
      <span><b>${winRate()}%</b>勝率</span>
      <span><b>${unlocked.length}</b>新報酬</span>`;
    goTo("end");
  }

  function renderAll() {
    renderBoard(); renderHud(); renderAchievements(); renderCollection(); renderSettings(); renderThreats(); renderHero();
  }

  function renderBoard() {
    const oldest = { player: state.queues.player[0], cpu: state.queues.cpu[0] };
    [...dom.board.children].forEach((cell, i) => {
      const owner = state.board[i];
      cell.className = `cell ${!owner && state.phase === "playing" && state.turn === "player" ? "playable" : "disabled"}`;
      cell.disabled = !!owner || state.phase !== "playing" || state.turn !== "player";
      cell.innerHTML = "";
      if (owner) {
        const piece = document.createElement("div"); piece.className = `piece ${owner}`;
        if (oldest[owner] === i && state.queues[owner].length >= 3) piece.classList.add("oldest");
        cell.appendChild(piece);
        if (oldest[owner] === i && state.queues[owner].length >= 3) cell.appendChild(Object.assign(document.createElement("div"), { className: "warning-ring" }));
      }
    });
  }

  function renderHud() {
    dom.turnChip.textContent = state.turn === "player" ? "YOUR TURN" : "CPU TURN";
    dom.turnChip.classList.toggle("cpu", state.turn === "cpu");
    dom.turnNumber.textContent = state.turnNumber;
    dom.winRate.textContent = `${winRate()}%`;
    dom.playerScore.textContent = state.save.stats.playerScore;
    dom.cpuScore.textContent = state.save.stats.cpuScore;
    dom.streak.textContent = state.save.stats.streak;
    dom.totalGames.textContent = state.save.stats.totalGames;
    const ach = achievementCount();
    dom.achievementProgressText.textContent = `${ach}/6`;
    dom.achievementProgressBar.style.width = `${ach / 6 * 100}%`;
    const col = collectionCount();
    dom.collectionProgressText.textContent = `${col}/${COLLECTION_TOTAL}`;
    dom.collectionProgressBar.style.width = `${Math.min(100, col / COLLECTION_TOTAL * 100)}%`;
    dom.miniAchievements.innerHTML = ACHIEVEMENTS.map(a => `<div class="mini-ach ${state.save.achievements[a.id] ? "unlocked" : ""} ${state.save.seenNew[a.id] ? "new" : ""}" title="${a.title}">${state.save.achievements[a.id] ? a.icon : "?"}</div>`).join("");
    dom.settingsSnap.innerHTML = `難易度: <b>${state.save.settings.difficulty}</b><br>エフェクト: <b>${state.save.settings.effects.toUpperCase()}</b><br>次に消える駒: <b>${oldestText()}</b>`;
  }

  function renderHero() {
    dom.heroWinRate.textContent = `${winRate()}%`;
    dom.heroGames.textContent = state.save.stats.totalGames;
    dom.heroCollection.textContent = `${Math.round(collectionCount() / COLLECTION_TOTAL * 100)}%`;
    dom.nextRewardHint.textContent = nextRewardHint();
  }

  function renderThreats() {
    const pos = snapshot();
    const playerWins = immediateWinningMoves(pos, "player");
    const cpuWins = immediateWinningMoves(pos, "cpu");
    const playerForks = countForks(pos, "player");
    const vanish = vanishTimingScore(pos, "player");
    if (vanish >= 26) state.roundFlags.vanishOpportunity = true;
    dom.threatPanel.innerHTML = [
      ["自分の即勝利", playerWins.length ? `${playerWins.map(n => n + 1).join(", ")}番が勝ち筋` : "なし"],
      ["CPUリーチ", cpuWins.length ? `${cpuWins.map(n => n + 1).join(", ")}番を警戒` : "安定"],
      ["フォーク圧", playerForks >= 2 ? "ダブルリーチ成立圏内" : "構築中"],
      ["消滅読み", vanish >= 26 ? "相手最古駒が防御を外す" : oldestText()]
    ].map(([a,b]) => `<div class="threat-card"><b>${a}</b><span>${b}</span></div>`).join("");
  }

  function renderAchievements() {
    const count = achievementCount();
    dom.achievementSummary.textContent = `解除率 ${Math.round(count / 6 * 100)}%`;
    dom.achievementGrid.innerHTML = ACHIEVEMENTS.map(a => {
      const unlocked = !!state.save.achievements[a.id];
      return `<article class="achievement-card ${unlocked ? "" : "locked"} ${state.save.seenNew[a.id] ? "new" : ""}"><div class="achievement-icon">${unlocked ? a.icon : "◼"}</div><h3>${a.title}</h3><p>${unlocked ? a.desc : "未解除: シルエット解析中。条件を満たすと開放されます。"}</p><div class="progress"><span style="width:${unlocked ? 100 : progressForAchievement(a.id)}%"></span></div></article>`;
    }).join("");
  }

  function renderCollection() {
    const entries = Object.entries(state.save.collection);
    dom.collectionSummary.textContent = `発見率 ${Math.round(collectionCount() / COLLECTION_TOTAL * 100)}%`;
    const cards = entries.map(([key, data], idx) => collectionCard(key, data, idx));
    for (let i = entries.length; i < COLLECTION_TOTAL; i++) cards.push(lockedCollectionCard(i));
    dom.collectionGrid.innerHTML = cards.join("");
  }

  function collectionCard(key, data, idx) {
    return `<article class="collection-card ${idx >= 5 ? "rare" : ""} ${state.save.seenNew[`col:${key}`] ? "new" : ""}">${miniBoard(data.board)}<h3>Pattern ${String(idx + 1).padStart(2,"0")}</h3><p>回転・反転同一判定済み / ${new Date(data.at).toLocaleDateString()}</p></article>`;
  }
  function lockedCollectionCard(i) { return `<article class="collection-card locked"><div class="collection-mini-board">${Array.from({length:9},()=>"<div class='mini-cell'>?</div>").join("")}</div><h3>Unknown ${String(i + 1).padStart(2,"0")}</h3><p>あと1戦で見つかるかもしれない未発見パターン。</p></article>`; }
  function miniBoard(board) { return `<div class="collection-mini-board">${board.map(v => `<div class="mini-cell ${v === "p" || v === "P" ? "p" : v === "c" || v === "C" ? "c" : ""}">${v === "p" || v === "P" ? "X" : v === "c" || v === "C" ? "O" : ""}</div>`).join("")}</div>`; }

  function renderSettings() {
    dom.bgmVolume.value = state.save.settings.bgmVolume;
    dom.seVolume.value = state.save.settings.seVolume;
    dom.difficultySelect.value = state.save.settings.difficulty;
    dom.effectsSelect.value = state.save.settings.effects;
    dom.colorAssist.checked = state.save.settings.colorAssist;
    dom.shakeEnabled.checked = state.save.settings.shake;
  }

  function showWinLine(line, winner) {
    const meta = LINE_META[line.join(",")]; if (!meta) return;
    Object.assign(dom.winLine.style, { top: meta.top, left: meta.left, width: meta.width, transform: `rotate(${meta.rotate}deg)` });
    dom.winLine.className = `win-line ${winner === "cpu" ? "cpu" : ""} show`;
  }

  function spawnPlacementEffect(index, owner) {
    const cell = dom.board.children[index];
    if (!cell || state.save.settings.effects === "low") return;
    for (let i = 0; i < (state.save.settings.effects === "high" ? 16 : 9); i++) {
      const p = document.createElement("i"); p.className = `particle ${owner === "cpu" ? "cpuP" : ""}`;
      const angle = Math.random() * Math.PI * 2, dist = 34 + Math.random() * 42;
      p.style.setProperty("--x", `${Math.cos(angle) * dist}px`); p.style.setProperty("--y", `${Math.sin(angle) * dist}px`);
      p.style.left = "50%"; p.style.top = "50%"; cell.appendChild(p); setTimeout(() => p.remove(), 620);
    }
  }
  function spawnVanishEffect(index, owner) { const cell = dom.board.children[index]; if (!cell) return; const r = document.createElement("i"); r.className = "ripple"; cell.appendChild(r); spawnPlacementEffect(index, owner); setTimeout(() => r.remove(), 620); }
  function flash() { dom.screenFlash.classList.remove("show"); void dom.screenFlash.offsetWidth; dom.screenFlash.classList.add("show"); }
  function pushFeed(text) { const item = document.createElement("div"); item.className = "feed-item"; item.textContent = text; dom.microFeed.prepend(item); while (dom.microFeed.children.length > 4) dom.microFeed.lastChild.remove(); }
  function showToast(title, body, achievement) { const t = document.createElement("div"); t.className = `toast ${achievement ? "achievement" : ""}`; t.innerHTML = `<b>${title}</b><br><span>${body}</span>`; dom.toastStack.appendChild(t); setTimeout(() => t.remove(), 4200); }
  function showHint() { playSound("ui"); const hints = ["自分が3駒ある時、次の配置でどの防御駒が消えるかを先に見ましょう。", "CPUの最古駒がラインを塞いでいるなら、あえて妨害せず消滅後の勝ち筋を作れます。", "フォークは2つの即勝利マスを同時に作る状態。HARD CPUも最優先で警戒します。", state.lastHint]; showToast("TACTICAL HINT", randomItem(hints), false); }

  function progressForAchievement(id) {
    if (id === "firstWin") return state.save.stats.wins ? 100 : 0;
    if (id === "streak3") return Math.min(100, state.save.stats.streak / 3 * 100);
    if (id === "complete") return achievementCount() / 5 * 100;
    return state.save.achievements[id] ? 100 : 35;
  }
  function achievementCount() { return ACHIEVEMENTS.filter(a => state.save.achievements[a.id]).length; }
  function collectionCount() { return Object.keys(state.save.collection).length; }
  function winRate() { return state.save.stats.totalGames ? Math.round(state.save.stats.wins / state.save.stats.totalGames * 100) : 0; }
  function oldestText() { const p = state.queues.player[0], c = state.queues.cpu[0]; return `YOU ${p == null ? "-" : p + 1} / CPU ${c == null ? "-" : c + 1}`; }
  function nextRewardHint() {
    if (!state.save.achievements.firstWin) return "次は『初勝利』を解除できます。中央と角から攻めましょう。";
    if (!state.save.achievements.streak3) return `3連勝まであと${Math.max(0, 3 - state.save.stats.streak)}勝。勝てば報酬が近づきます。`;
    if (collectionCount() < COLLECTION_TOTAL) return `勝利図鑑コンプリートまであと${COLLECTION_TOTAL - collectionCount()}種類。NEWパターンを狙いましょう。`;
    return "全データが高水準。HARDで勝率80%を目指しましょう。";
  }

  function unlockAudio() { startAmbientTone(); }
  function startAmbientTone() {
    if (state.audio) return;
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const bgmGain = ctx.createGain(); bgmGain.gain.value = state.save.settings.muted ? 0 : state.save.settings.bgmVolume; bgmGain.connect(ctx.destination);
      const osc = ctx.createOscillator(); osc.type = "sine"; osc.frequency.value = 74;
      const filter = ctx.createBiquadFilter(); filter.type = "lowpass"; filter.frequency.value = 260;
      osc.connect(filter); filter.connect(bgmGain); osc.start(); state.audio = { ctx, bgmGain };
    } catch { state.audio = { disabled: true }; }
  }
  function playSound(type) {
    if (state.save.settings.muted) return;
    startAmbientTone(); if (!state.audio || state.audio.disabled) return;
    const ctx = state.audio.ctx, gain = ctx.createGain(), osc = ctx.createOscillator();
    const vol = state.save.settings.seVolume; gain.gain.value = 0.0001; gain.connect(ctx.destination); osc.connect(gain);
    const map = { placePlayer:[330,620,.09,"triangle"], placeCpu:[220,440,.12,"sine"], vanish:[140,60,.18,"sawtooth"], win:[440,880,.42,"triangle"], lose:[180,70,.36,"sine"], ui:[520,620,.06,"sine"], achievement:[660,1320,.5,"triangle"] };
    const [from,to,dur,wave] = map[type] || map.ui; osc.type = wave; osc.frequency.setValueAtTime(from, ctx.currentTime); osc.frequency.exponentialRampToValueAtTime(Math.max(20,to), ctx.currentTime + dur); gain.gain.exponentialRampToValueAtTime(Math.max(.001, vol * .22), ctx.currentTime + .018); gain.gain.exponentialRampToValueAtTime(.0001, ctx.currentTime + dur); osc.start(); osc.stop(ctx.currentTime + dur + .03);
  }
})();
