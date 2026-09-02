(() => {
  const NativeAudioContext = window.AudioContext || window.webkitAudioContext;
  if (!NativeAudioContext) return;

  let sharedContext = null;

  const getSharedContext = () => {
    if (!sharedContext || sharedContext.state === 'closed') {
      sharedContext = new NativeAudioContext();
    }
    return sharedContext;
  };

  const unlockAudio = () => {
    const context = getSharedContext();
    const resume = context.state === 'suspended' ? context.resume() : Promise.resolve();

    Promise.resolve(resume).then(() => {
      // Een vrijwel stille buffer binnen de gebruikersactie maakt mobiele Safari/Chrome
      // duidelijk dat audio door de gebruiker is geactiveerd.
      const buffer = context.createBuffer(1, 1, context.sampleRate || 44100);
      const source = context.createBufferSource();
      source.buffer = buffer;
      source.connect(context.destination);
      source.start(0);
    }).catch(() => {
      // Een volgende tik probeert opnieuw te ontgrendelen.
    });
  };

  function SharedAudioContext() {
    return getSharedContext();
  }

  SharedAudioContext.prototype = NativeAudioContext.prototype;
  try { Object.setPrototypeOf(SharedAudioContext, NativeAudioContext); } catch (_) {}

  window.AudioContext = SharedAudioContext;
  window.webkitAudioContext = SharedAudioContext;

  // Capture-fase: ontgrendel vóór de Start-knop zijn submit-handler uitvoert.
  document.addEventListener('pointerdown', unlockAudio, { capture: true, passive: true });
  document.addEventListener('touchstart', unlockAudio, { capture: true, passive: true });
})();
