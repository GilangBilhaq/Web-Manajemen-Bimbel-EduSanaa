(function () {
  var view = window.PORTAL_VIEW || 'tutor';
  var base = window.APPS_SCRIPT_BASE || '';
  var src = base + (base.indexOf('?') === -1 ? '?' : '&') + 'view=' + encodeURIComponent(view);

  var frame = document.getElementById('app');
  var loader = document.getElementById('loader');

  function hideLoader() { if (loader) loader.classList.add('hidden'); }

  // Kalau lambat / diblokir browser, tampilkan link cadangan untuk buka langsung
  var fallbackTimer = setTimeout(function () {
    if (!loader || loader.classList.contains('hidden')) return;
    var fb = document.createElement('div');
    fb.className = 'loader-fallback show';
    fb.innerHTML = 'Masih memuat... koneksi lambat?<br><a href="' + src + '" target="_top" rel="noopener">Buka langsung &rarr;</a>';
    loader.appendChild(fb);
  }, 12000);

  frame.addEventListener('load', function () {
    clearTimeout(fallbackTimer);
    setTimeout(hideLoader, 300);
  });

  frame.src = src;
})();
