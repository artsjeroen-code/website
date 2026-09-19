(() => {
  const THEME_KEY = 'startpagina-theme';
  const TASKS_KEY = 'focus-tasks-v1';
  const BLOCKS_KEY = 'focus-blocks-v1';
  const SOUND_KEY = 'focus-sound-v1';

  const durations = { focus: 25 * 60, short: 5 * 60, long: 15 * 60 };
  const labels = {
    focus: ['Focus', 'Tijd voor ongestoord werken'],
    short: ['Korte pauze', 'Even opladen'],
    long: ['Lange pauze', 'Tijd voor een langere onderbreking']
  };

  let mode = 'focus';
  let totalSeconds = durations.focus;
  let remainingSeconds = totalSeconds;
  let running = false;
  let timerId = null;
  let endAt = null;
  let blocks = Math.max(0, Math.min(4, Number(localStorage.getItem(BLOCKS_KEY)) || 0));
  let tasks = loadTasks();
  let activeTaskId = tasks.find(t => t.active && !t.completed)?.id || null;
  let soundEnabled = localStorage.getItem(SOUND_KEY) !== 'false';

  const $ = id => document.getElementById(id);
  const timerDisplay = $('timerDisplay');
  const timerState = $('timerState');
  const timerRing = $('timerRing');
  const startPauseButton = $('startPauseButton');
  const taskDoneButton = $('taskDoneButton');
  const taskList = $('taskList');
  const emptyState = $('emptyState');
  const taskInput = $('taskInput');
  const soundToggle = $('soundToggle');

  function loadTasks() {
    try {
      const parsed = JSON.parse(localStorage.getItem(TASKS_KEY) || '[]');
      return Array.isArray(parsed) ? parsed : [];
    } catch {
      return [];
    }
  }

  function saveTasks() {
    localStorage.setItem(TASKS_KEY, JSON.stringify(tasks));
  }

  function formatTime(seconds) {
    const safe = Math.max(0, Math.ceil(seconds));
    const minutes = Math.floor(safe / 60);
    const secs = safe % 60;
    return `${String(minutes).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }

  function renderTimer() {
    timerDisplay.textContent = formatTime(remainingSeconds);
    const elapsed = Math.max(0, totalSeconds - remainingSeconds);
    const progress = totalSeconds > 0 ? (elapsed / totalSeconds) * 360 : 0;
    timerRing.style.setProperty('--progress', `${Math.min(360, progress)}deg`);
    document.title = running ? `${formatTime(remainingSeconds)} · Focus Timer` : 'Focus Timer';
  }

  function renderMode() {
    $('modeLabel').textContent = labels[mode][0];
    $('modeDescription').textContent = labels[mode][1];
    document.querySelectorAll('.mode-button').forEach(button => {
      button.classList.toggle('active', button.dataset.mode === mode);
    });
  }

  function renderBlocks() {
    document.querySelectorAll('#progressDots span').forEach((dot, index) => {
      dot.classList.toggle('done', index < blocks);
    });
    $('blockCount').textContent = `${blocks} / 4 focusblokken`;
  }

  function renderSound() {
    soundToggle.setAttribute('aria-pressed', String(soundEnabled));
    soundToggle.setAttribute('aria-label', soundEnabled ? 'Geluid uitschakelen' : 'Geluid inschakelen');
    soundToggle.querySelector('span').textContent = soundEnabled ? '🔊' : '🔇';
  }

  function renderTasks() {
    taskList.replaceChildren();
    emptyState.hidden = tasks.length > 0;

    tasks.forEach(task => {
      const li = document.createElement('li');
      li.className = 'task-item';
      if (task.completed) li.classList.add('completed');
      if (task.id === activeTaskId && !task.completed) li.classList.add('active');

      const select = document.createElement('button');
      select.type = 'button';
      select.className = 'task-select';
      select.setAttribute('aria-label', task.completed ? 'Taak opnieuw openen' : 'Taak selecteren');
      select.addEventListener('click', () => {
        if (task.completed) {
          task.completed = false;
          activeTaskId = task.id;
        } else {
          activeTaskId = task.id;
        }
        tasks.forEach(item => { item.active = item.id === activeTaskId; });
        saveTasks();
        renderTasks();
      });

      const text = document.createElement('span');
      text.className = 'task-text';
      text.textContent = task.text;
      text.tabIndex = 0;
      text.addEventListener('click', () => select.click());
      text.addEventListener('keydown', event => {
        if (event.key === 'Enter' || event.key === ' ') {
          event.preventDefault();
          select.click();
        }
      });

      const remove = document.createElement('button');
      remove.type = 'button';
      remove.className = 'task-delete';
      remove.textContent = '×';
      remove.setAttribute('aria-label', `Verwijder ${task.text}`);
      remove.addEventListener('click', () => {
        tasks = tasks.filter(item => item.id !== task.id);
        if (activeTaskId === task.id) activeTaskId = null;
        saveTasks();
        renderTasks();
      });

      li.append(select, text, remove);
      taskList.appendChild(li);
    });

    taskDoneButton.disabled = !activeTaskId;
  }

  function addTask(text) {
    const clean = text.trim();
    if (!clean) return;
    const task = {
      id: crypto.randomUUID ? crypto.randomUUID() : String(Date.now() + Math.random()),
      text: clean,
      completed: false,
      active: false
    };
    tasks.unshift(task);
    if (!activeTaskId) {
      activeTaskId = task.id;
      task.active = true;
    }
    saveTasks();
    renderTasks();
  }

  function finishActiveTask() {
    if (!activeTaskId) return;
    const task = tasks.find(item => item.id === activeTaskId);
    if (!task) return;
    task.completed = true;
    task.active = false;
    activeTaskId = tasks.find(item => !item.completed && item.id !== task.id)?.id || null;
    tasks.forEach(item => { item.active = item.id === activeTaskId; });
    saveTasks();
    renderTasks();
    playTone(660, .12);
  }

  function setMode(nextMode, preserveRunning = false) {
    mode = nextMode;
    totalSeconds = durations[mode];
    remainingSeconds = totalSeconds;
    endAt = preserveRunning ? Date.now() + remainingSeconds * 1000 : null;
    if (!preserveRunning) stopTimer();
    renderMode();
    renderTimer();
    timerState.textContent = 'Klaar om te starten';
  }

  function startTimer() {
    if (running) return;
    running = true;
    endAt = Date.now() + remainingSeconds * 1000;
    startPauseButton.textContent = 'Pauze';
    timerState.textContent = mode === 'focus' ? 'Bezig met focus' : 'Pauze loopt';
    timerId = window.setInterval(tick, 250);
    tick();
  }

  function pauseTimer() {
    if (!running) return;
    tick();
    running = false;
    window.clearInterval(timerId);
    timerId = null;
    endAt = null;
    startPauseButton.textContent = 'Verder';
    timerState.textContent = 'Gepauzeerd';
  }

  function stopTimer() {
    running = false;
    if (timerId) window.clearInterval(timerId);
    timerId = null;
    endAt = null;
    startPauseButton.textContent = 'Start';
  }

  function tick() {
    if (!running || !endAt) return;
    remainingSeconds = Math.max(0, (endAt - Date.now()) / 1000);
    renderTimer();
    if (remainingSeconds <= 0) completeSession();
  }

  function completeSession() {
    stopTimer();
    remainingSeconds = 0;
    renderTimer();
    playTone(784, .18);

    if (mode === 'focus') {
      blocks += 1;
      if (blocks >= 4) {
        blocks = 0;
        localStorage.setItem(BLOCKS_KEY, String(blocks));
        renderBlocks();
        setMode('long');
        timerState.textContent = 'Focusblok klaar · lange pauze';
      } else {
        localStorage.setItem(BLOCKS_KEY, String(blocks));
        renderBlocks();
        setMode('short');
        timerState.textContent = 'Focusblok klaar · korte pauze';
      }
    } else {
      setMode('focus');
      timerState.textContent = 'Pauze klaar · tijd voor focus';
    }
  }

  function addMinutes(minutes) {
    const seconds = minutes * 60;
    remainingSeconds += seconds;
    totalSeconds += seconds;
    if (running && endAt) endAt += seconds * 1000;
    renderTimer();
    timerState.textContent = `+${minutes} minuten toegevoegd`;
  }

  function resetTimer() {
    stopTimer();
    totalSeconds = durations[mode];
    remainingSeconds = totalSeconds;
    timerState.textContent = 'Klaar om te starten';
    renderTimer();
  }

  function skipTimer() {
    stopTimer();
    if (mode === 'focus') {
      setMode(blocks === 3 ? 'long' : 'short');
    } else {
      setMode('focus');
    }
  }

  function playTone(frequency, seconds) {
    if (!soundEnabled) return;
    try {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (!AudioContext) return;
      const context = new AudioContext();
      const oscillator = context.createOscillator();
      const gain = context.createGain();
      oscillator.frequency.value = frequency;
      gain.gain.setValueAtTime(.08, context.currentTime);
      gain.gain.exponentialRampToValueAtTime(.001, context.currentTime + seconds);
      oscillator.connect(gain);
      gain.connect(context.destination);
      oscillator.start();
      oscillator.stop(context.currentTime + seconds);
      oscillator.addEventListener('ended', () => context.close());
    } catch {}
  }

  $('railTheme')?.addEventListener('click', () => {
    const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
    document.documentElement.dataset.theme = next;
    localStorage.setItem(THEME_KEY, next);
  });

  soundToggle.addEventListener('click', () => {
    soundEnabled = !soundEnabled;
    localStorage.setItem(SOUND_KEY, String(soundEnabled));
    renderSound();
    if (soundEnabled) playTone(660, .08);
  });

  document.querySelectorAll('.mode-button').forEach(button => {
    button.addEventListener('click', () => setMode(button.dataset.mode));
  });

  startPauseButton.addEventListener('click', () => running ? pauseTimer() : startTimer());
  $('resetButton').addEventListener('click', resetTimer);
  $('skipButton').addEventListener('click', skipTimer);
  taskDoneButton.addEventListener('click', finishActiveTask);

  document.querySelectorAll('[data-add-minutes]').forEach(button => {
    button.addEventListener('click', () => addMinutes(Number(button.dataset.addMinutes)));
  });

  $('taskForm').addEventListener('submit', event => {
    event.preventDefault();
    addTask(taskInput.value);
    taskInput.value = '';
    taskInput.focus();
  });

  renderMode();
  renderTimer();
  renderBlocks();
  renderTasks();
  renderSound();
})();
