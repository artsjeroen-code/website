(() => {
  'use strict';

  if (window.location.hostname === 'test-artsjeroen.ddns.net') {
    document.documentElement.classList.add('staging-environment');
  }
})();
