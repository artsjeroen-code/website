(() => {
  const STORAGE_KEY = 'ademhaling-settings-v1';
  const DEFAULTS_VERSION = 2;

  const form = document.getElementById('settingsForm');
  const inhaleInput = document.getElementById('inhale');
  const holdFullInput = document.getElementById('holdFull');
  const exhaleInput = document.getElementById('exhale');
  const holdEmptyInput = document.getElementById('holdEmpty');
  const repetitionsInput = document.getElementById('repetitions');
  const breathingPreset = document.getElementById('breathingPreset');
  const presetNote = document.getElementById('presetNote');
  const soundEnabled = document.getElementById('soundEnabled');
  const visualMode = document.getElementById('visualMode');

  const phaseLabel = document.getElementById('phaseLabel');
  const phaseTime = document.getElementById('phaseTime');
  const ball = document.getElementById('breathingBall');
  const ballSpace = document.getElementById('ballSpace');
  const breathingGraph = document.getElementById('breathingGraph');
  const breathingPath = document.getElementById('breathingPath');
  const breathingMarker = document.getElementById('breathingMarker');
  const cycleDuration = document.getElementById('cycleDuration');
  const totalDuration = document.getElementById('totalDuration');
  const remainingDuration = document.getElementById('remainingDuration');
  const repetitionStatus = document.getElementById('repetitionStatus');
  const validationMessage = document.getElementById('validationMessage');

  const startButton = document.getElementById('startButton');
  const pauseButton = document.getElementById('pauseButton');
  const resetButton = document.getElementById('resetButton');

  const inputs = [inhaleInput, holdFullInput, exhaleInput, holdEmptyInput, repetitionsInput];
  const phaseInputs = [inhaleInput, holdFullInput, exhaleInput, holdEmptyInput];

  const presets = {
    '4-7-8-0': {
      values: [4, 7, 8, 0],
      note: 'Diepe ontspanning — rustig patroon met lange uitademing; vaak gebruikt om tot rust te komen.'
    },
    '4-4-4-4': {
      values: [4, 4, 4, 4],
      note: 'Focus & kalmte — box breathing met vier gelijke fasen voor een strak, aandachtig ritme.'
    },
    '4-0-8-0': {
      values: [4, 0, 8, 0],
      note: 'Ontstressen — eenvoudige langzame ademhaling met extra nadruk op de uitademing.'
    },
    '5-0-5-0': {
      values: [5, 0, 5, 0],
      note: 'Rustig ritme — zes ademhalingen per minuut; gelijkmatig en geschikt voor rustige dagelijkse oefening.'
    },
    '7-0-11-0': {
      values: [7, 0, 11, 0],
      note: 'Diepe rust — zeer langzaam ritme met lange uitademing; kies dit alleen als het comfortabel voelt.'
    },
    '4-4-4-0': {
      values: [4, 4, 4, 0],
      note: 'Driehoek-focus — inademen, vasthouden en uitademen in drie gelijke delen, zonder slotpauze.'
    }
  };

  let state = 'idle';
  let animationFrame = null;
  let session = null;
  let startedAt = 0;
  let pausedAt = 0;
  let totalPaused = 0;
  let audioContext = null;
  let lastPhaseToken = '';

  const clampInt = (value, min, max) => {
    const parsed = Number.parseInt(value, 10);
    if (!Number.isFinite(parsed)) return min;
    return Math.min(max, Math.max(min, parsed));
  };

  const readSettings = () => {
    const settings = {
      inhale: clampInt(inhaleInput.value, 0, 20),
      holdFull: clampInt(holdFullInput.value, 0, 20),
      exhale: clampInt(exhaleInput.value, 0, 20),
      holdEmpty: clampInt(holdEmptyInput.value, 0, 20),
      repetitions: clampInt(repetitionsInput.value, 1, 99)
    };

    inhaleInput.value = settings.inhale;
    holdFullInput.value = settings.holdFull;
    exhaleInput.value = settings.exhale;
    holdEmptyInput.value = settings.holdEmpty;
    repetitionsInput.value = settings.repetitions;
    return settings;
  };

  const savePreferences = () => {
    const settings = readSettings();
    const payload = {
      defaultsVersion: DEFAULTS_VERSION,
      inhale: settings.inhale,
      holdFull: settings.holdFull,
      exhale: settings.exhale,
      holdEmpty: settings.holdEmpty,
      repetitions: settings.repetitions,
      preset: breathingPreset.value,
      sound: Boolean(soundEnabled.checked),
      visual: visualMode.value === 'ball' ? 'ball' : 'graph'
    };

    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (_) {
      // De tool blijft werken als lokale opslag niet beschikbaar is.
    }
  };

  const restorePreferences = () => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      inhaleInput.value = clampInt(saved.inhale, 0, 20);
      holdFullInput.value = clampInt(saved.holdFull, 0, 20);
      exhaleInput.value = clampInt(saved.exhale, 0, 20);
      holdEmptyInput.value = clampInt(saved.holdEmpty, 0, 20);
      repetitionsInput.value = saved.defaultsVersion === DEFAULTS_VERSION
        ? clampInt(saved.repetitions, 1, 99)
        : 10;
      soundEnabled.checked = saved.sound !== false;
      visualMode.value = saved.visual === 'ball' ? 'ball' : 'graph';
    } catch (_) {
      // Ongeldige of geblokkeerde opslag valt terug op de HTML-standaardwaarden.
    }
  };

  const formatClock = (seconds) => {
    const total = Math.max(0, Math.ceil(seconds));
    const minutes = Math.floor(total / 60);
    const remainder = total % 60;
    return `${String(minutes).padStart(2, '0')}:${String(remainder).padStart(2, '0')}`;
  };

  const formatWholeSeconds = (seconds) => String(Math.max(0, Math.ceil(seconds)));

  const ensureAudio = () => {
    if (!soundEnabled.checked) return null;
    const AudioContextClass = window.AudioContext || window.webkitAudioContext;
    if (!AudioContextClass) return null;
    if (!audioContext) audioContext = new AudioContextClass();
    if (audioContext.state === 'suspended') audioContext.resume();
    return audioContext;
  };

  const playTone = (frequency = 392, duration = 0.34, volume = 0.012) => {
    if (!soundEnabled.checked) return;
    const ctx = ensureAudio();
    if (!ctx) return;

    const oscillator = ctx.createOscillator();
    const gain = ctx.createGain();
    const filter = ctx.createBiquadFilter();
    const now = ctx.currentTime;

    oscillator.type = 'sine';
    oscillator.frequency.setValueAtTime(frequency, now);
    filter.type = 'lowpass';
    filter.frequency.setValueAtTime(1100, now);
    filter.Q.setValueAtTime(0.5, now);

    gain.gain.setValueAtTime(0.0001, now);
    gain.gain.linearRampToValueAtTime(volume, now + 0.07);
    gain.gain.exponentialRampToValueAtTime(0.0001, now + duration);

    oscillator.connect(filter);
    filter.connect(gain);
    gain.connect(ctx.destination);
    oscillator.start(now);
    oscillator.stop(now + duration + 0.04);
  };

  const phaseTone = (key) => {
    const tones = {
      inhale: 440.0,
      holdFull: 493.88,
      exhale: 369.99,
      holdEmpty: 329.63
    };
    playTone(tones[key] || 392, 0.34, 0.011);
  };

  const getCycleSeconds = (settings) => settings.inhale + settings.holdFull + settings.exhale + settings.holdEmpty;

  const makePhases = (settings) => [
    { key: 'inhale', label: 'Inademen', duration: settings.inhale, from: 1, to: 2 },
    { key: 'holdFull', label: 'Vasthouden', duration: settings.holdFull, from: 2, to: 2 },
    { key: 'exhale', label: 'Uitademen', duration: settings.exhale, from: 2, to: 1 },
    { key: 'holdEmpty', label: 'Vasthouden', duration: settings.holdEmpty, from: 1, to: 1 }
  ];

  const easeInOut = (progress) => (1 - Math.cos(Math.PI * Math.min(1, Math.max(0, progress)))) / 2;

  const scaleForPhase = (phase, elapsed) => {
    if (!phase || phase.duration <= 0) return phase ? phase.to : 1;
    if (phase.from === phase.to) return phase.from;
    const progress = easeInOut(elapsed / phase.duration);
    return phase.from + (phase.to - phase.from) * progress;
  };

  const scaleToGraphY = (scale) => 205 - (Math.min(2, Math.max(1, scale)) - 1) * 150;

  const phaseAtElapsed = (phases, elapsedInCycle) => {
    let cursor = 0;
    for (const phase of phases) {
      if (phase.duration <= 0) continue;
      const end = cursor + phase.duration;
      if (elapsedInCycle <= end) {
        return { phase, elapsed: Math.min(phase.duration, Math.max(0, elapsedInCycle - cursor)) };
      }
      cursor = end;
    }
    const fallback = [...phases].reverse().find((phase) => phase.duration > 0);
    return fallback ? { phase: fallback, elapsed: fallback.duration } : null;
  };

  const updateGraphPath = (settings = readSettings()) => {
    const cycle = getCycleSeconds(settings);
    if (cycle <= 0) {
      breathingPath.setAttribute('d', '');
      breathingMarker.setAttribute('cx', '0');
      breathingMarker.setAttribute('cy', '205');
      return;
    }

    const phases = makePhases(settings);
    const points = [];
    const samples = 160;
    for (let i = 0; i <= samples; i += 1) {
      const elapsed = cycle * (i / samples);
      const located = phaseAtElapsed(phases, elapsed);
      const scale = located ? scaleForPhase(located.phase, located.elapsed) : 1;
      const x = 1000 * (i / samples);
      const y = scaleToGraphY(scale);
      points.push(`${i === 0 ? 'M' : 'L'} ${x.toFixed(2)} ${y.toFixed(2)}`);
    }
    breathingPath.setAttribute('d', points.join(' '));
  };

  const setGraphMarker = (elapsedInCycle, cycleSeconds, scale) => {
    const x = cycleSeconds > 0 ? (elapsedInCycle / cycleSeconds) * 1000 : 0;
    breathingMarker.setAttribute('cx', String(Math.min(1000, Math.max(0, x))));
    breathingMarker.setAttribute('cy', String(scaleToGraphY(scale)));
  };

  const applyVisualMode = () => {
    const graphMode = visualMode.value !== 'ball';
    breathingGraph.classList.toggle('is-hidden', !graphMode);
    ballSpace.classList.toggle('is-hidden', graphMode);
    if (graphMode) updateGraphPath();
  };

  const syncPresetFromInputs = () => {
    const key = phaseInputs.map((input) => clampInt(input.value, 0, 20)).join('-');
    if (presets[key]) {
      breathingPreset.value = key;
      presetNote.textContent = presets[key].note;
    } else {
      breathingPreset.value = 'custom';
      presetNote.textContent = 'Eigen ritme — pas de vier fasen vrij aan tussen 0 en 20 seconden.';
    }
  };

  const applyPreset = () => {
    const preset = presets[breathingPreset.value];
    if (!preset) {
      presetNote.textContent = 'Eigen ritme — pas de vier fasen vrij aan tussen 0 en 20 seconden.';
      savePreferences();
      return;
    }

    preset.values.forEach((value, index) => {
      phaseInputs[index].value = value;
    });
    presetNote.textContent = preset.note;
    updateSummary();
    updateGraphPath();
    savePreferences();
  };

  const updateSummary = () => {
    const settings = readSettings();
    const cycle = getCycleSeconds(settings);
    const total = cycle * settings.repetitions;

    cycleDuration.textContent = formatClock(cycle);
    totalDuration.textContent = formatClock(total);

    if (state === 'idle' || state === 'finished') {
      remainingDuration.textContent = formatClock(total);
      repetitionStatus.textContent = `0 / ${settings.repetitions}`;
    }

    validationMessage.textContent = cycle === 0 ? 'Kies minimaal één fase met een duur groter dan 0 seconden.' : '';
  };

  const buildSession = (settings) => {
    const cycleSeconds = getCycleSeconds(settings);
    return {
      settings,
      phases: makePhases(settings),
      cycleSeconds,
      totalSeconds: cycleSeconds * settings.repetitions
    };
  };

  const setInputsDisabled = (disabled) => {
    inputs.forEach((input) => { input.disabled = disabled; });
    breathingPreset.disabled = disabled;
  };

  const setBallScale = (scale) => { ball.style.transform = `scale(${scale})`; };

  const locatePhase = (elapsedInCycle) => {
    const located = phaseAtElapsed(session.phases, elapsedInCycle);
    if (!located) return null;
    return { ...located.phase, elapsed: located.elapsed };
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
    if (!currentPhase) return finishSession();

    const token = `${repetitionIndex}:${currentPhase.key}`;
    if (state === 'running' && token !== lastPhaseToken) {
      lastPhaseToken = token;
      phaseTone(currentPhase.key);
    }

    const scale = scaleForPhase(currentPhase, currentPhase.elapsed);
    const phaseRemaining = currentPhase.duration - currentPhase.elapsed;
    const totalRemaining = session.totalSeconds - elapsed;

    setBallScale(scale);
    setGraphMarker(elapsedInCycle, session.cycleSeconds, scale);
    phaseLabel.textContent = currentPhase.label;
    phaseTime.textContent = formatWholeSeconds(phaseRemaining);
    remainingDuration.textContent = formatClock(totalRemaining);
    repetitionStatus.textContent = `${repetitionIndex + 1} / ${session.settings.repetitions}`;
  };

  const tick = (now) => {
    if (state !== 'running') return;
    renderRunningState(now);
    if (state === 'running') animationFrame = requestAnimationFrame(tick);
  };

  const startSession = () => {
    const settings = readSettings();
    const cycleSeconds = getCycleSeconds(settings);
    if (cycleSeconds <= 0) {
      validationMessage.textContent = 'Kies minimaal één fase met een duur groter dan 0 seconden.';
      return;
    }

    savePreferences();
    ensureAudio();
    session = buildSession(settings);
    updateGraphPath(settings);
    state = 'running';
    startedAt = performance.now();
    pausedAt = 0;
    totalPaused = 0;
    lastPhaseToken = '';

    setInputsDisabled(true);
    startButton.disabled = true;
    pauseButton.disabled = false;
    pauseButton.textContent = 'Pauze';
    validationMessage.textContent = '';
    repetitionStatus.textContent = `1 / ${settings.repetitions}`;
    remainingDuration.textContent = formatClock(session.totalSeconds);

    cancelAnimationFrame(animationFrame);
    renderRunningState(startedAt);
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
    playTone(523.25, 0.55, 0.014);
    setBallScale(1);
    setGraphMarker(0, 1, 1);
    phaseLabel.textContent = 'Klaar';
    phaseTime.textContent = '0';
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
    lastPhaseToken = '';

    setInputsDisabled(false);
    setBallScale(1);
    setGraphMarker(0, 1, 1);
    phaseLabel.textContent = 'Klaar';
    phaseTime.textContent = '0';
    startButton.disabled = false;
    startButton.textContent = 'Start';
    pauseButton.disabled = true;
    pauseButton.textContent = 'Pauze';
    syncPresetFromInputs();
    updateSummary();
    updateGraphPath();
    savePreferences();
  };

  form.addEventListener('submit', (event) => {
    event.preventDefault();
    startSession();
  });

  breathingPreset.addEventListener('change', applyPreset);
  soundEnabled.addEventListener('change', savePreferences);
  visualMode.addEventListener('change', () => {
    applyVisualMode();
    savePreferences();
  });
  pauseButton.addEventListener('click', pauseSession);
  resetButton.addEventListener('click', resetSession);

  inputs.forEach((input) => input.addEventListener('input', () => {
    if (state === 'idle' || state === 'finished') {
      if (phaseInputs.includes(input)) syncPresetFromInputs();
      updateSummary();
      updateGraphPath();
      savePreferences();
    }
  }));

  restorePreferences();
  syncPresetFromInputs();
  applyVisualMode();
  updateSummary();
  updateGraphPath();
  setGraphMarker(0, 1, 1);
})();
