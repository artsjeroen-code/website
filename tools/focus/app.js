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
  const soundEnabledInput = $('soundEnabled');

  function loadTasks() {
    try {
      const parsed = JSON.parse(localStorage.getItem(TASKS_KEY) || '[]');
      if (!Array.isArray(parsed)) return [];
      return parsed.map(task => ({
        ...task,
        note: typeof task.note === 'string' ? task.note : ''
      }));
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
    const progressRatio = totalSeconds > 0 ? elapsed / totalSeconds : 0;
    const progressDegrees = Math.min(360, progressRatio * 360);
    const progressPercent = Math.min(100, progressRatio * 100);
    timerRing.style.setProperty('--progress', progressDegrees + 'deg');
    timerRing.style.setProperty('--progress-percent', progressPercent + '%');
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
    soundEnabledInput.checked = soundEnabled;
    soundEnabledInput.setAttribute('aria-label', soundEnabled ? 'Geluid uitschakelen' : 'Geluid inschakelen');
  }

  function closeTaskOverlays(exceptItem = null) {
    document.querySelectorAll('.task-item').forEach(item => {
      if (item === exceptItem) return;
      item.querySelector('.task-menu')?.setAttribute('hidden', '');
      item.querySelector('.task-info-wrap')?.classList.remove('show-note');
    });
  }

  function setTaskCompletion(task, completed) {
    task.completed = completed;
    task.active = false;

    if (completed && activeTaskId === task.id) {
      activeTaskId = tasks.find(item => !item.completed && item.id !== task.id)?.id || null;
    } else if (!completed) {
      activeTaskId = task.id;
    }

    tasks.forEach(item => { item.active = item.id === activeTaskId; });
    saveTasks();
    renderTasks();
  }

  function moveTask(taskId, direction) {
    const index = tasks.findIndex(item => item.id === taskId);
    if (index < 0) return;

    let target = index;
    if (direction === 'up') target = Math.max(0, index - 1);
    if (direction === 'down') target = Math.min(tasks.length - 1, index + 1);
    if (direction === 'top') target = 0;
    if (direction === 'bottom') target = tasks.length - 1;
    if (target === index) return;

    const [task] = tasks.splice(index, 1);
    tasks.splice(target, 0, task);
    saveTasks();
    renderTasks();
  }

  function renderTasks() {
    taskList.replaceChildren();
    emptyState.hidden = tasks.length > 0;

    tasks.forEach(task => {
      const li = document.createElement('li');
      li.className = 'task-item';
      if (task.completed) li.classList.add('completed');
      if (task.id === activeTaskId && !task.completed) li.classList.add('active');

      const main = document.createElement('div');
      main.className = 'task-main';

      const select = document.createElement('button');
      select.type = 'button';
      select.className = 'task-select';
      select.setAttribute('aria-label', task.completed ? 'Afgeronde taak' : 'Taak selecteren');
      select.addEventListener('click', () => {
        if (task.completed) return;
        activeTaskId = task.id;
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

      const actions = document.createElement('div');
      actions.className = 'task-actions';

      const infoWrap = document.createElement('div');
      infoWrap.className = 'task-info-wrap';

      const info = document.createElement('button');
      info.type = 'button';
      info.className = 'task-info';
      info.textContent = 'i';
      info.setAttribute('aria-label', task.note ? `Notitie bij ${task.text}` : `Geen notitie bij ${task.text}`);
      info.setAttribute('aria-expanded', 'false');
      if (!task.note) info.classList.add('is-empty');

      if (task.note) {
        const bubble = document.createElement('div');
        bubble.className = 'task-note-bubble';
        bubble.setAttribute('role', 'tooltip');
        bubble.textContent = task.note;
        infoWrap.append(info, bubble);

        info.addEventListener('click', event => {
          event.stopPropagation();
          const show = !infoWrap.classList.contains('show-note');
          closeTaskOverlays(show ? li : null);
          infoWrap.classList.toggle('show-note', show);
          info.setAttribute('aria-expanded', String(show));
        });
      } else {
        infoWrap.append(info);
      }

      const menuButton = document.createElement('button');
      menuButton.type = 'button';
      menuButton.className = 'task-more';
      menuButton.textContent = '⋯';
      menuButton.setAttribute('aria-label', `Opties voor ${task.text}`);
      menuButton.setAttribute('aria-expanded', 'false');

      const menu = document.createElement('div');
      menu.className = 'task-menu';
      menu.setAttribute('hidden', '');

      const makeMenuButton = (label, action, className = '') => {
        const button = document.createElement('button');
        button.type = 'button';
        button.textContent = label;
        if (className) button.className = className;
        button.addEventListener('click', event => {
          event.stopPropagation();
          action();
        });
        return button;
      };

      menu.append(
        makeMenuButton(task.completed ? 'Opnieuw openen' : 'Taak afronden', () => {
          setTaskCompletion(task, !task.completed);
        })
      );

      const moveLabel = document.createElement('span');
      moveLabel.className = 'task-menu-label';
      moveLabel.textContent = 'Verplaatsen';
      menu.append(moveLabel);

      const moveGrid = document.createElement('div');
      moveGrid.className = 'task-move-grid';
      moveGrid.append(
        makeMenuButton('Omhoog', () => moveTask(task.id, 'up')),
        makeMenuButton('Omlaag', () => moveTask(task.id, 'down')),
        makeMenuButton('Bovenaan', () => moveTask(task.id, 'top')),
        makeMenuButton('Onderaan', () => moveTask(task.id, 'bottom'))
      );
      menu.append(moveGrid);

      const noteAction = makeMenuButton(task.note ? 'Notitie bewerken' : 'Notitie toevoegen', () => {
        menu.setAttribute('hidden', '');
        menuButton.setAttribute('aria-expanded', 'false');
        editor.removeAttribute('hidden');
        textarea.value = task.note || '';
        textarea.focus();
        textarea.setSelectionRange(textarea.value.length, textarea.value.length);
      });
      menu.append(noteAction);

      menu.append(
        makeMenuButton('Verwijderen', () => {
          tasks = tasks.filter(item => item.id !== task.id);
          if (activeTaskId === task.id) {
            activeTaskId = tasks.find(item => !item.completed)?.id || null;
            tasks.forEach(item => { item.active = item.id === activeTaskId; });
          }
          saveTasks();
          renderTasks();
        }, 'danger')
      );

      menuButton.addEventListener('click', event => {
        event.stopPropagation();
        const opening = menu.hasAttribute('hidden');
        closeTaskOverlays(opening ? li : null);
        if (opening) {
          menu.removeAttribute('hidden');
        } else {
          menu.setAttribute('hidden', '');
        }
        menuButton.setAttribute('aria-expanded', String(opening));
      });

      menu.addEventListener('click', event => event.stopPropagation());

      actions.append(infoWrap, menuButton, menu);
      main.append(select, text, actions);

      const editor = document.createElement('div');
      editor.className = 'task-note-editor';
      editor.setAttribute('hidden', '');

      const textarea = document.createElement('textarea');
      textarea.maxLength = 500;
      textarea.rows = 3;
      textarea.placeholder = 'Korte notitie bij deze taak…';
      textarea.setAttribute('aria-label', `Notitie bij ${task.text}`);

      const editorActions = document.createElement('div');
      editorActions.className = 'task-note-editor-actions';

      const cancelNote = document.createElement('button');
      cancelNote.type = 'button';
      cancelNote.textContent = 'Annuleren';
      cancelNote.addEventListener('click', () => {
        editor.setAttribute('hidden', '');
      });

      const saveNote = document.createElement('button');
      saveNote.type = 'button';
      saveNote.className = 'primary-small';
      saveNote.textContent = 'Opslaan';
      saveNote.addEventListener('click', () => {
        task.note = textarea.value.trim();
        saveTasks();
        renderTasks();
      });

      editorActions.append(cancelNote, saveNote);
      editor.append(textarea, editorActions);

      li.append(main, editor);
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
      active: false,
      note: ''
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

  soundEnabledInput.addEventListener('change', () => {
    soundEnabled = soundEnabledInput.checked;
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

  document.addEventListener('click', () => closeTaskOverlays());

  renderMode();
  renderTimer();
  renderBlocks();
  renderTasks();
  renderSound();
})();
