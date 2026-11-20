(function () {
  document.documentElement.classList.add('js');

  // Filter, search and sort the prospect table
  var tbody = document.querySelector('#prospects tbody');
  if (tbody) {
    var rows = [].slice.call(tbody.rows);
    var q = document.getElementById('q'), ft = document.getElementById('f-type'), fr = document.getElementById('f-tier'),
        fw = document.getElementById('f-town'), fs = document.getElementById('f-sort'),
        count = document.getElementById('count'), empty = document.getElementById('empty');
    var sorters = {
      score: function (a, b) { return b.dataset.score - a.dataset.score || a.dataset.miles - b.dataset.miles; },
      miles: function (a, b) { return a.dataset.miles - b.dataset.miles || b.dataset.score - a.dataset.score; },
      name: function (a, b) { return a.dataset.name.localeCompare(b.dataset.name); }
    };
    function apply() {
      var words = q.value.toLowerCase().trim().split(/\s+/).filter(Boolean), shown = 0;
      rows.sort(sorters[fs.value] || sorters.score).forEach(function (r) {
        var ok = (!ft.value || r.dataset.type === ft.value) && (!fr.value || r.dataset.tier === fr.value) &&
                 (!fw.value || r.dataset.town === fw.value) &&
                 words.every(function (w) { return r.dataset.text.indexOf(w) > -1; });
        r.hidden = !ok; if (ok) shown++;
        tbody.appendChild(r);
      });
      count.textContent = shown; empty.hidden = shown > 0;
    }
    [q, ft, fr, fw, fs].forEach(function (el) { el.addEventListener('input', apply); el.addEventListener('change', apply); });
  }

  // Fade-ins
  var els = [].slice.call(document.querySelectorAll('.reveal'));
  if (document.documentElement.classList.contains('anim') && 'IntersectionObserver' in window) {
    var io = new IntersectionObserver(function (es) {
      es.forEach(function (e) { if (e.isIntersecting) { e.target.classList.add('is-in'); io.unobserve(e.target); } });
    }, { rootMargin: '0px 0px -8% 0px' });
    els.forEach(function (el, i) { el.style.transitionDelay = (i % 4) * 60 + 'ms'; io.observe(el); });
  } else els.forEach(function (el) { el.classList.add('is-in'); });

  // Highlight the section in the nav
  var links = [].slice.call(document.querySelectorAll('.nav a'));
  if ('IntersectionObserver' in window && links.length) {
    var map = {}; links.forEach(function (a) { map[a.getAttribute('href').slice(1)] = a; });
    var so = new IntersectionObserver(function (es) {
      es.forEach(function (e) { if (e.isIntersecting) { links.forEach(function (a) { a.classList.remove('is-on'); }); if (map[e.target.id]) map[e.target.id].classList.add('is-on'); } });
    }, { rootMargin: '-45% 0px -50% 0px' });
    Object.keys(map).forEach(function (id) { var s = document.getElementById(id); if (s) so.observe(s); });
  }
})();
