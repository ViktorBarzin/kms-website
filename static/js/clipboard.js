(function () {
  "use strict";

  function flash(el, label) {
    var orig = el.dataset.orig || el.textContent;
    el.dataset.orig = orig;
    el.classList.add("copied");
    if (el.tagName === "BUTTON") el.textContent = label;
    setTimeout(function () {
      el.classList.remove("copied");
      if (el.tagName === "BUTTON") el.textContent = orig;
    }, 1200);
  }

  async function copyText(t) {
    if (navigator.clipboard && window.isSecureContext) {
      try { await navigator.clipboard.writeText(t); return true; } catch (e) { /* fall through */ }
    }
    var ta = document.createElement("textarea");
    ta.value = t;
    ta.style.position = "fixed";
    ta.style.opacity = "0";
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    var ok = false;
    try { ok = document.execCommand("copy"); } catch (e) { ok = false; }
    document.body.removeChild(ta);
    return ok;
  }

  // 1. inline <code class="copy" data-copy="...">
  document.querySelectorAll("code.copy[data-copy]").forEach(function (el) {
    el.addEventListener("click", async function () {
      var ok = await copyText(el.dataset.copy);
      if (ok) flash(el, "Copied");
    });
  });

  // 2. button.btn-copy[data-copy-target=elementId]
  document.querySelectorAll(".btn-copy[data-copy-target]").forEach(function (btn) {
    btn.addEventListener("click", async function () {
      var target = document.getElementById(btn.dataset.copyTarget);
      if (!target) return;
      var text = target.innerText.trim();
      var ok = await copyText(text);
      if (ok) flash(btn, "Copied!");
    });
  });
})();
