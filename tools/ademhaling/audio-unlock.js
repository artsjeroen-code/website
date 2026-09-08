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
      const buffer = context.createBuffer(1, 1, context.sampleRate || 44100);
      const source = context.createBufferSource();
      source.buffer = buffer;
      source.connect(context.destination);
      source.start(0);
    }).catch(() => {});
  };

  function SharedAudioContext() { return getSharedContext(); }
  SharedAudioContext.prototype = NativeAudioContext.prototype;
  try { Object.setPrototypeOf(SharedAudioContext, NativeAudioContext); } catch (_) {}

  window.AudioContext = SharedAudioContext;
  window.webkitAudioContext = SharedAudioContext;

  document.addEventListener('pointerdown', unlockAudio, { capture: true, passive: true });
  document.addEventListener('touchstart', unlockAudio, { capture: true, passive: true });
})();
