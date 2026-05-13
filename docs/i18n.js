(function () {
  var STORAGE_KEY = 'flextext-lang';
  var SUPPORTED = ['en', 'id'];
  var DEFAULT = 'en';
  var cache = {};
  var current = DEFAULT;

  function detectLang() {
    var stored = localStorage.getItem(STORAGE_KEY);
    if (stored && SUPPORTED.indexOf(stored) !== -1) return stored;
    var nav = (navigator.language || '').toLowerCase();
    return nav.startsWith('id') ? 'id' : DEFAULT;
  }

  function resolvePath(key, obj) {
    return key.split('.').reduce(function (o, k) {
      return (o != null && o[k] !== undefined) ? o[k] : undefined;
    }, obj);
  }

  function updateButton() {
    var code = document.getElementById('lang-code');
    if (code) code.textContent = current.toUpperCase();
  }

  function closeDropdown() {
    var dd = document.getElementById('lang-dropdown');
    if (dd) dd.hidden = true;
  }

  function apply(t) {
    document.querySelectorAll('[data-i18n]').forEach(function (el) {
      var v = resolvePath(el.getAttribute('data-i18n'), t);
      if (v !== undefined) el.textContent = v;
    });
    document.querySelectorAll('[data-i18n-html]').forEach(function (el) {
      var v = resolvePath(el.getAttribute('data-i18n-html'), t);
      if (v !== undefined) el.innerHTML = v;
    });
    document.documentElement.lang = current;
    updateButton();
  }

  function loadAndApply(lang) {
    if (cache[lang]) {
      apply(cache[lang]);
      return;
    }
    fetch('locales/' + lang + '.json')
      .then(function (r) { return r.json(); })
      .then(function (t) {
        cache[lang] = t;
        apply(t);
      })
      .catch(function () {
        console.warn('FlexText i18n: could not load locales/' + lang + '.json');
      });
  }

  function setLang(lang) {
    if (SUPPORTED.indexOf(lang) === -1) return;
    current = lang;
    localStorage.setItem(STORAGE_KEY, lang);
    closeDropdown();
    loadAndApply(lang);
  }

  function toggleDropdown() {
    var dd = document.getElementById('lang-dropdown');
    if (!dd) return;
    dd.hidden = !dd.hidden;
  }

  window.i18n = {
    setLang: setLang,
    toggleDropdown: toggleDropdown,
    toggleLang: function () { setLang(current === 'en' ? 'id' : 'en'); }
  };

  function init() {
    current = detectLang();
    if (current === DEFAULT) {
      updateButton();
    } else {
      loadAndApply(current);
    }

    document.addEventListener('click', function (e) {
      var switcher = document.getElementById('lang-switcher');
      if (switcher && !switcher.contains(e.target)) {
        closeDropdown();
      }
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
