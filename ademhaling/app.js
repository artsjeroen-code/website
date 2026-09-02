(() => {
  const form = document.getElementById('settingsForm');
  const inhaleInput = document.getElementById('inhale');
  const holdFullInput = document.getElementById('holdFull');
  const exhaleInput = document.getElementById('exhale');
  const holdEmptyInput = document.getElementById('holdEmpty');
  const repetitionsInput = document.getElementById('repetitions');

  const phaseLabel = document.getElementById('phaseLabel');
  const phaseTime = document.getElementById('phaseTime');
  const ball = document.getElementById('breathingBall');
  const cycleDuration = document.getElementById('cycleDuration');
  const totalDuration = document.getElementById('totalDuration');
  const remainingDuration = document.getElementById('remainingDuration');
  const repetitionStatus = document.getElementById('repetitionStatus');
  const validationMessage = document.getElementById('validationMessage');

  const startButton = document.getElementById('startButton');
  const pauseButton = document.getElementById('pauseButton');
  const resetButton = document.getElementById('resetButton');

  const inputs = [inhaleInput, holdFullInput, exhaleInput, holdEmptyInput, repetitionsInput];

  let state = 'idle';
  let animationFrame = null;
  let session = null;
  let startedAt = 0;
  let pausedAt = 0;
  let totalPaused = 0;

  const clampInt = (value, min, max) => {
    const parsed = Number.parseInt(value, 10);
    if (!Number.isFinite(parsed)) return min;
    return Math.min(max, Math.max(min, parsed));
  };

  const formatClock = (seconds) => {
    const total = Math.max(0, Math.ceil(seconds));
    const minutes = Math.floor(total / 60);
    const remainder = total % 60;
    return `${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}`;
  };

  const formatDecimal = (seconds) => `${Math.max(0, seconds).toFixed(1).replace('.', ',')} s`;

  const readSettings = () => {
    const settings = {
      inhale: clampInt(inhaleInput.value, 0, 9),
      holdFull: clampInt(holdFullInput.value, 0, 9),
      exhale: clampInt(exhaleInput.value, 0, 9),
      holdEmpty: clampInt(holdEmptyInput.value, 0, 9),
      repetitions: clampInt(repetitionsInput.value, 1, 99)
    };

    inhaleInput.value = settings.inhale;
    holdFullInput.value = settings.holdFull;
    exhaleInput.value = settings.exhale;
    holdEmptyInput.value = settings.holdEmpty;
    repetitionsInput.value = settings.repetitions;

    return settings;
  };

  const getCycleSeconds = (settings) => settings.inhale + settings.holdFull + settings.exhale + settings.holdEmpty;

  const updateSummary = () => {
    const settings = readSettings();
    const cycle = getCycleSeconds(settings);
    const total = cycle * settings.repetitions;

    cycleDuration.textContent = formatClock(cycle);
    totalDuration.textContent = formatClock(total);

    if (state === 'idle') {
      remainingDuration.textContent = formatClock(total);
      repetitionStatus.textContent = `0 / ${settings.repetitions}`;
    }

    validationMessage.textContent = cycle === 0 ? 'Kies minimaal één fase met een duur groter dan 0 seconden.' : '';
  };

  const buildSession = (settings) => {
    const phases = [
      { key: 'inhale', label: 'Inademen', duration: settings.inhale, from: 1, to: 2 },
      { key: 'holdFull', label: 'Vasthouden', duration: settings.holdFull, from: 2, to: 2 },
      { key: 'exhale', label: 'Uitademen', duration: settings.exhale, from: 2, to: 1 },
      { key: 'holdEmpty', label: 'Vasthouden', duration: settings.holdEmpty, from: 1, to: 1 }
    ];

    const cycleSeconds = getCycleSeconds(settings);
    return {
      settings,
      phases,
      cycleSeconds,
      totalSeconds: cycleSeconds * settings.repetitions
    };
  };

  const setInputsDisabled = (disabled) => inputs.forEach((input) => { input.disabled = disabled; });

  const setBallScale = (scale) => {
    ball.style.transform = `scale(${scale})`;
  };

  const locatePhase = (elapsedInCycle) => {
    let cursor = 0;

    for (const phase of session.phases) {
      if (phase.duration <= 0) continue;
      const end = cursor + phase.duration;
      if (elapsedInCycle < end || Math.abs(elapsedInCycle - end) < 0.000001) {
        return {
          ...phase,
          elapsed: Math.min(phase.duration, Math.max(0, elapsedInCycle - cursor))
        };
      }
      cursor = end;
    }

    const fallback = [...session.phases].reverse().find((phase) => phase.duration > 0);
    return fallback ? { ...fallback, elapsed: fallback.duration } : null;
  };

  const renderRunningState = (now) => {
    if (!session) return;

    const activeNow = state === 'paused' ? pausedAt : now;
    const elapsed = Math.max(0, (activeNow - startedAt - totalPaused) / 1000);

    if (elapsed >= session.totalSeconds) {
      finishSession();
      return;
    }

    const repetitionIndex = Math.floor(elapsed / session.cycleSeconds);
    const elapsedInCycle = elapsed - repetitionIndex * session.cycleSeconds;
    const currentPhase = locatePhase(elapsedInCycle);

    if (!currentPhase) {
      finishSession();
      return;
    }

    const progress = currentPhase.duration > 0 ? Math.min(1, currentPhase.elapsed / currentPhase.duration) : 1;
    const scale = currentPhase.from + (currentPhase.to - currentPhase.from) * progress;
    const phaseRemaining = currentPhase.duration - currentPhase.elapsed;
    const totalRemaining = session.totalSeconds - elapsed;

    setBallScale(scale);
    phaseLabel.textContent = currentPhase.label;
    phaseTime.textContent = formatDecimal(phaseRemaining);
    remainingDuration.textContent = formatClock(totalRemaining);
    repetitionStatus.textContent = `${repetitionIndex + 1} / ${session.settings.repetitions}`;
  };

  const tick = (now) => {
    if (state !== 'running') return;
    renderRunningState(now);
    animationFrame = requestAnimationFrame(tick);
  };

  const startSession = () => {
    const settings = readSettings();
    const cycleSeconds = getCycleSeconds(settings);

    if (cycleSeconds <= 0) {
      validationMessage.textContent = 'Kies minimaal één fase met een duur groter dan 0 seconden.';
      return;
    }

    session = buildSession(settings);
    state = 'running';
    startedAt = performance.now();
    pausedAt = 0;
    totalPaused = 0;

    setInputsDisabled(true);
    startButton.disabled = true;
    pauseButton.disabled = false;
    pauseButton.textContent = 'Pauze';
    validationMessage.textContent = '';
    repetitionStatus.textContent = `1 / ${settings.repetitions}`;
    remainingDuration.textContent = formatClock(session.totalSeconds);

    cancelAnimationFrame(animationFrame);
    animationFrame = requestAnimationFrame(tick);
  };

  const pauseSession = () => {
    if (state === 'running') {
      state = 'paused';
      pausedAt = performance.now();
      cancelAnimationFrame(animationFrame);
      renderRunningState(pausedAt);
      pauseButton.textContent = 'Hervat';
      phaseLabel.textContent += ' — gepauzeerd';
      return;
    }

    if (state === 'paused') {
      const resumedAt = performance.now();
      totalPaused += resumedAt - pausedAt;
      pausedAt = 0;
      state = 'running';
      pauseButton.textContent = 'Pauze';
      animationFrame = requestAnimationFrame(tick);
    }
  };

  const finishSession = () => {
    cancelAnimationFrame(animationFrame);
    state = 'finished';
    setBallScale(1);
    phaseLabel.textContent = 'Klaar';
    phaseTime.textContent = '0,0 s';
    remainingDuration.textContent = '00:00';
    repetitionStatus.textContent = `${session.settings.repetitions} / ${session.settings.repetitions}`;
    pauseButton.disabled = true;
    pauseButton.textContent = 'Pauze';
    startButton.disabled = false;
    startButton.textContent = 'Opnieuw';
    setInputsDisabled(false);
  };

  const resetSession = () => {
    cancelAnimationFrame(animationFrame);
    state = 'idle';
    session = null;
    startedAt = 0;
    pausedAt = 0;
    totalPaused = 0;

    setInputsDisabled(false);
    setBallScale(1);
    phaseLabel.textContent = 'Klaar';
    phaseTime.textContent = '0,0 s';
    startButton.disabled = false;
    startButton.textContent = 'Start';
    pauseButton.disabled = true;
    pauseButton.textContent = 'Pauze';
    updateSummary();
  };

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    startSession();
  });

  pauseButton.addEventListener('click', pauseSession);
  resetButton.addEventListener('click', resetSession);

  inputs.forEach((input) => input.addEventListener('input', () => {
    if (state === 'idle' || state === 'finished') updateSummary();
  }));

  updateSummary();
})();
