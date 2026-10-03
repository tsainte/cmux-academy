// Shared helpers: storage, theme, nav, progress. Data comes from assets/data.js (window.ACADEMY).
(function () {
  const A = window.ACADEMY;
  const PREFIX = "cmuxacademy.";

  const store = {
    get(key, fallback) {
      try {
        const raw = localStorage.getItem(PREFIX + key);
        return raw === null ? fallback : JSON.parse(raw);
      } catch (e) { return fallback; }
    },
    set(key, value) {
      try { localStorage.setItem(PREFIX + key, JSON.stringify(value)); } catch (e) { /* storage unavailable */ }
    },
  };

  const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  // `code` spans in data strings become <code>
  const rich = (s) => esc(s).replace(/`([^`]+)`/g, "<code>$1</code>")
    .replace(/(https:\/\/[^\s<)]+[^\s<).,])/g, '<a href="$1" target="_blank" rel="noopener">$1</a>');

  const moduleById = Object.fromEntries(A.modules.map((m) => [m.id, m]));

  function applyTheme() {
    const t = store.get("theme", null);
    if (t) document.documentElement.setAttribute("data-theme", t);
  }

  function renderNav(active) {
    const links = [
      ["index.html", "Home", "home"],
      ["slides.html", "Slides", "slides"],
      ["tasks.html", "Tasks", "tasks"],
      ["drill.html", "Drill", "drill"],
      ["reference.html", "Reference", "reference"],
      ["reading.html", "Reading", "reading"],
    ];
    const el = document.getElementById("site-header");
    if (!el) return;
    el.innerHTML =
      '<header class="site"><div class="wrap">' +
      '<a class="brand" href="index.html">cmux<span>/</span>academy</a>' +
      '<nav class="main" aria-label="Main">' +
      links.map(([href, label, key]) => `<a href="${href}"${key === active ? ' aria-current="page"' : ""}>${label}</a>`).join("") +
      "</nav>" +
      '<button class="theme-btn" id="theme-btn" type="button" aria-label="Toggle light or dark theme">Theme</button>' +
      "</div></header>";
    document.getElementById("theme-btn").addEventListener("click", () => {
      const cur = document.documentElement.getAttribute("data-theme") ||
        (matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light");
      const next = cur === "dark" ? "light" : "dark";
      document.documentElement.setAttribute("data-theme", next);
      store.set("theme", next);
    });
  }

  const doneSet = () => new Set(store.get("done", []));
  function setDone(id, on) {
    const s = doneSet();
    on ? s.add(id) : s.delete(id);
    store.set("done", [...s]);
  }
  function progress(moduleId) {
    const done = doneSet();
    const tasks = A.tasks.filter((t) => !moduleId || t.m === moduleId);
    const n = tasks.filter((t) => done.has(t.id)).length;
    return { done: n, total: tasks.length, pct: tasks.length ? Math.round((100 * n) / tasks.length) : 0 };
  }

  window.Academy = { A, store, esc, rich, moduleById, applyTheme, renderNav, doneSet, setDone, progress };
  applyTheme();
})();
