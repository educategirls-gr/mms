// The state this website serves. Every page takes the name from here: the tab
// title, the headings, the sidebar and the footers. There is a copy of EG-MMS
// per state now (Rajasthan since 1 Oct 2026), so each says which one it is.
// A new state's copy changes this one line, plus STATE_NAME and SITE_HOST at
// the top of Code.gs, and nothing else.
var STATE_NAME = 'Uttar Pradesh';

(function () {
  // <title data-state-title="EG MMS | {STATE} Analytics Portal">: the plain
  // title in the tag is what shows if this file ever fails to load.
  var t = document.querySelector('title[data-state-title]');
  if (t) document.title = t.getAttribute('data-state-title').replace('{STATE}', STATE_NAME);

  // <span class="state-name"></span> anywhere in the page gets the name.
  function fill() {
    var els = document.querySelectorAll('.state-name');
    for (var i = 0; i < els.length; i++) els[i].textContent = STATE_NAME;
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', fill);
  else fill();
})();
