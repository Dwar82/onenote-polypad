/* Polypad for OneNote — task pane logic
 * Personal, non-commercial use. Uses Polypad's free JavaScript API.
 */

(function () {
  "use strict";

  let pp = null;            // Polypad instance
  let isOneNote = false;    // true when running inside OneNote

  const $ = (id) => document.getElementById(id);

  function setStatus(text) { $("status").textContent = text; }

  let toastTimer = null;
  function toast(msg, ms) {
    const t = $("toast");
    t.textContent = msg;
    t.classList.remove("hidden");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add("hidden"), ms || 2600);
  }

  function enableButtons(on) {
    ["btnInsert", "btnDownload", "btnSave", "btnLoad", "btnClear"].forEach((id) => {
      $(id).disabled = !on;
    });
    // "Insert into page" only makes sense inside OneNote
    $("btnInsert").disabled = !on || !isOneNote;
    if (!isOneNote) $("btnInsert").title = "Only available when running inside OneNote";
  }

  // ---- Polypad setup ---------------------------------------------------------
  function initPolypad() {
    // Show the full manipulative library. Omit apiKey for the free tier.
    pp = Polypad.create($("polypad"), {
      // apiKey: "",           // add a Mathigon-issued key here only if you get one
      sidebarSettings: true,
      toolbar: true,
      settings: true,
      initial: { options: { grid: "square-grid", canvas: "infinite" } }
    });

    if (pp.bindKeyboardEvents) pp.bindKeyboardEvents();
    if (Polypad.loadFonts) Polypad.loadFonts();

    // Keep Polypad sized to the pane
    window.addEventListener("resize", () => pp && pp.resize && pp.resize());
  }

  // ---- Get a PNG data URI of the current canvas -----------------------------
  async function getPng(width, height) {
    // PolypadInstance.image(width, height, type) -> Promise<dataURI>
    return await pp.image(width || 1400, height || 900, "png");
  }

  // ---- Insert snapshot onto the active OneNote page -------------------------
  async function insertSnapshot() {
    if (!isOneNote) { toast("Not running inside OneNote."); return; }
    setStatus("Inserting…");
    try {
      const dataUrl = await getPng(1400, 900);
      await OneNote.run(async (context) => {
        const page = context.application.getActivePage();
        // addOutline(left, top, html) — drop the snapshot as an image outline
        page.addOutline(40, 40,
          '<p><img src="' + dataUrl + '" alt="Polypad snapshot" width="600" /></p>');
        await context.sync();
      });
      setStatus("Ready");
      toast("Snapshot added to your page.");
    } catch (err) {
      console.error(err);
      setStatus("Ready");
      toast("Couldn't insert directly — use “Download PNG” and drag it in.", 4200);
    }
  }

  // ---- Download the snapshot as a PNG file (reliable fallback) ---------------
  async function downloadPng() {
    setStatus("Rendering…");
    try {
      const dataUrl = await getPng(1600, 1000);
      const a = document.createElement("a");
      a.href = dataUrl;
      a.download = "polypad-" + Date.now() + ".png";
      document.body.appendChild(a);
      a.click();
      a.remove();
      toast("PNG downloaded — drag it onto your page.");
    } catch (err) {
      console.error(err);
      toast("Could not render image.");
    }
    setStatus("Ready");
  }

  // ---- Save / load canvas as JSON -------------------------------------------
  function openSavePanel() {
    const data = pp.serialize();
    $("jsonTitle").textContent = "Save canvas";
    $("jsonHint").textContent = "Copy this text and paste it into a OneNote page (or anywhere) to reopen this canvas later.";
    $("jsonBox").value = JSON.stringify(data);
    $("jsonBox").readOnly = true;
    $("jsonApply").classList.add("hidden");
    $("jsonCopy").classList.remove("hidden");
    $("jsonPanel").classList.remove("hidden");
  }

  function openLoadPanel() {
    $("jsonTitle").textContent = "Load canvas";
    $("jsonHint").textContent = "Paste previously saved canvas text here, then choose “Load this canvas”.";
    $("jsonBox").value = "";
    $("jsonBox").readOnly = false;
    $("jsonApply").classList.remove("hidden");
    $("jsonCopy").classList.add("hidden");
    $("jsonPanel").classList.remove("hidden");
    $("jsonBox").focus();
  }

  function applyLoad() {
    const raw = $("jsonBox").value.trim();
    if (!raw) { toast("Nothing to load."); return; }
    try {
      const data = JSON.parse(raw);
      pp.unSerialize(data);
      $("jsonPanel").classList.add("hidden");
      toast("Canvas loaded.");
    } catch (err) {
      console.error(err);
      toast("That doesn't look like valid canvas data.");
    }
  }

  async function copyJson() {
    const text = $("jsonBox").value;
    try {
      await navigator.clipboard.writeText(text);
      toast("Copied to clipboard.");
    } catch (e) {
      $("jsonBox").select();
      document.execCommand("copy");
      toast("Copied.");
    }
  }

  function clearCanvas() {
    if (pp && pp.clear) { pp.clear(); toast("Canvas cleared (undo with Ctrl+Z)."); }
  }

  // ---- Wire everything up ----------------------------------------------------
  function wireUi() {
    $("btnInsert").addEventListener("click", insertSnapshot);
    $("btnDownload").addEventListener("click", downloadPng);
    $("btnSave").addEventListener("click", openSavePanel);
    $("btnLoad").addEventListener("click", openLoadPanel);
    $("btnClear").addEventListener("click", clearCanvas);
    $("jsonClose").addEventListener("click", () => $("jsonPanel").classList.add("hidden"));
    $("jsonCopy").addEventListener("click", copyJson);
    $("jsonApply").addEventListener("click", applyLoad);
  }

  Office.onReady((info) => {
    isOneNote = info && info.host === Office.HostType.OneNote;
    try {
      initPolypad();
      wireUi();
      enableButtons(true);
      setStatus(isOneNote ? "Ready" : "Ready (preview mode)");
    } catch (err) {
      console.error(err);
      setStatus("Failed to load Polypad");
      toast("Polypad failed to load. Check your internet connection.", 5000);
    }
  });
})();
