/* CMS Contract Generator — client-only
   - PIN gate: 0623 (sessionStorage)
   - Requires uploading: CSV + Buyer DOCX template + Seller DOCX template
   - Generates editable .docx contracts via docxtemplater/pizzip
   - Supports:
       * Lot-by-lot download (buyer only / seller only)
       * ZIP downloads: buyer-only, seller-only, or all (buyer+seller)
       * ZIP includes selected lots (or all if none selected)

   CSV “Parsing…” hardening:
   - Verifies Papa is loaded (otherwise shows clear error)
   - Try/catch around parse
   - worker:true to avoid UI lockup
   - 10s watchdog that errors if Papa never calls back
*/

(() => {
  const CONFIG = {
    PIN: "0623",

    // Required headers (exact match). We validate these.
    REQUIRED_COLS: [
      "Contract #",
      "Consignor",
      "Buyer",
      "Lot Number #2"
    ],

    // DOCX placeholder -> CSV column
    MAP: {
      contract_no: "Contract #",
      consignor: "Consignor",
      buyer: "Buyer",
      lot_no: "Lot Number #2",
      head_count: "Head Count",
      breed: "Breed",
      sex: "Sex",
      base_weight: "Base Weight",
      delivery: "Delivery",
      year: "Year",
      location: "Location",
      shrink: "Shrink",
      slide: "Slide",
      description: "Description",
      second_description: "Second Description",
      price_cwt: "Calculated High Bid",
      down_money_due: "Down Money Due"
    },

    // Optional fallbacks if a preferred column is empty
    FALLBACKS: {
      lot_no: ["Lot Number #2", "Lot Number"],
      price_cwt: ["Calculated High Bid"]
    }
  };

  // ---------- DOM ----------
  const $ = (id) => document.getElementById(id);

  const pinOverlay = $("pinOverlay");
  const pinInput = $("pinInput");
  const pinUnlockBtn = $("pinUnlockBtn");
  const pinClearBtn = $("pinClearBtn");
  const pinErr = $("pinErr");

  const sessionPill = $("sessionPill");
  const exitBtn = $("exitBtn");

  const csvInput = $("csvInput");
  const buyerTplInput = $("buyerTplInput");
  const sellerTplInput = $("sellerTplInput");

  const csvMeta = $("csvMeta");
  const buyerTplMeta = $("buyerTplMeta");
  const sellerTplMeta = $("sellerTplMeta");

  const dzCsv = $("dzCsv");
  const dzBuyerTpl = $("dzBuyerTpl");
  const dzSellerTpl = $("dzSellerTpl");

  const validationBox = $("validationBox");

  const generateBtn = $("generateBtn");
  const clearBtn = $("clearBtn");

  const zipBuyerBtn = $("zipBuyerBtn");
  const zipSellerBtn = $("zipSellerBtn");
  const zipAllBtn = $("zipAllBtn");

  const lotCount = $("lotCount");
  const resultsBox = $("resultsBox");

  const selectAll = $("selectAll");
  const onlyWithBuyer = $("onlyWithBuyer");
  const onlyWithConsignor = $("onlyWithConsignor");

  // ---------- State ----------
  const state = {
    csvFile: null,
    csvRowsRaw: [],
    buyerTplFile: null,
    sellerTplFile: null,
    buyerTplBuf: null,
    sellerTplBuf: null,
    lots: [] // generated
  };

  // ---------- UI helpers ----------
  function setStatus(type, msg, show = true) {
    if (!show) {
      validationBox.style.display = "none";
      validationBox.textContent = "";
      validationBox.className = "status";
      return;
    }
    validationBox.style.display = "block";
    validationBox.textContent = msg;
    validationBox.className = `status ${type || ""}`.trim();
  }

  function sanitizeFilename(name) {
    if (!name) return "Untitled";
    return String(name)
      .replace(/[\/\\:*?"<>|]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  }

  function escapeHtml(str) {
    return String(str || "")
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;")
      .replaceAll("'","&#039;");
  }

  function getCell(row, colName) {
    if (!row || !colName) return "";
    const v = row[colName];
    return (v === undefined || v === null) ? "" : String(v).trim();
  }

  function moneyClean(v) {
    if (v === undefined || v === null) return "";
    const s = String(v).trim();
    if (!s) return "";
    return s.replace(/\$/g, "").replace(/,/g, "").trim();
  }

  function requiredMissing(headers) {
    const set = new Set((headers || []).map(h => String(h).trim()));
    return CONFIG.REQUIRED_COLS.filter(req => !set.has(req));
  }

  function canGenerate() {
    return !!(state.csvRowsRaw.length && state.buyerTplBuf && state.sellerTplBuf);
  }

  function setGenerateEnabled() {
    generateBtn.disabled = !canGenerate();
  }

  function setDownloadEnabled(enabled) {
    zipBuyerBtn.disabled = !enabled;
    zipSellerBtn.disabled = !enabled;
    zipAllBtn.disabled = !enabled;
  }

  // ---------- PIN ----------
  function isUnlocked() {
    return sessionStorage.getItem("cms_pin_ok") === "1";
  }
  function lock() {
    sessionStorage.removeItem("cms_pin_ok");
    sessionPill.textContent = "Locked";
    pinOverlay.style.display = "flex";
    pinInput.value = "";
    pinErr.style.display = "none";
  }
  function unlock() {
    sessionStorage.setItem("cms_pin_ok", "1");
    sessionPill.textContent = "Unlocked";
    pinOverlay.style.display = "none";
  }

  pinUnlockBtn.addEventListener("click", () => {
    const v = (pinInput.value || "").trim();
    if (v === CONFIG.PIN) unlock();
    else pinErr.style.display = "block";
  });
  pinClearBtn.addEventListener("click", () => {
    pinInput.value = "";
    pinErr.style.display = "none";
    pinInput.focus();
  });
  pinInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") pinUnlockBtn.click();
  });

  exitBtn.addEventListener("click", () => {
    clearSessionFiles(true);
    lock();
  });

  // ---------- Drag/drop ----------
  function wireDropzone(zoneEl, inputEl, onFiles) {
    zoneEl.addEventListener("dragover", (e) => {
      e.preventDefault();
      zoneEl.style.borderColor = "rgba(51,102,153,.75)";
    });
    zoneEl.addEventListener("dragleave", () => {
      zoneEl.style.borderColor = "rgba(51,102,153,.35)";
    });
    zoneEl.addEventListener("drop", (e) => {
      e.preventDefault();
      zoneEl.style.borderColor = "rgba(51,102,153,.35)";
      const files = e.dataTransfer?.files;
      if (files && files.length) onFiles(files);
    });
    inputEl.addEventListener("change", () => {
      const files = inputEl.files;
      if (files && files.length) onFiles(files);
    });
  }

  // ---------- CSV ----------
  function assertLibsForCsv() {
    if (typeof Papa === "undefined") {
      throw new Error(
        "PapaParse is not loaded (Papa is undefined).\n\n" +
        "Fix:\n" +
        "1) In GitHub, confirm /lib/papa.min.js exists and is named EXACTLY that (case-sensitive).\n" +
        "2) On the live site, open DevTools → Network → refresh and confirm /lib/papa.min.js returns 200 (not 404).\n"
      );
    }
  }

  async function handleCsv(files) {
    const f = files[0];
    if (!f) return;

    state.csvFile = f;
    csvMeta.textContent = `${f.name} (${Math.round(f.size/1024)} KB)`;

    setStatus("warn", "Parsing CSV…");

    // Watchdog: if Papa never calls complete/error, we break out with a clear message.
    let watchdog = null;
    const WATCHDOG_MS = 10000;

    try {
      assertLibsForCsv();

      // Clear existing rows before parsing
      state.csvRowsRaw = [];
      setGenerateEnabled();
      setDownloadEnabled(false);
      renderLots([]);

      await new Promise((resolve, reject) => {
        let finished = false;

        watchdog = setTimeout(() => {
          if (!finished) {
            reject(new Error(
              "CSV parse timed out.\n\n" +
              "Most common causes:\n" +
              "- /lib/papa.min.js did not load (404/case mismatch)\n" +
              "- A JS error occurred before Papa completed\n" +
              "- Opening the page as a local file instead of GitHub Pages\n\n" +
              "Check DevTools Console + Network."
            ));
          }
        }, WATCHDOG_MS);

        Papa.parse(f, {
          header: true,
          skipEmptyLines: true,
          worker: true,
          complete: (res) => {
            finished = true;
            clearTimeout(watchdog);

            const data = res.data || [];
            const headers = res.meta?.fields || Object.keys(data[0] || {});
            const missing = requiredMissing(headers);

            state.csvRowsRaw = data;

            if (!data.length) {
              setStatus("bad", "CSV parsed, but there are zero rows.");
              state.csvRowsRaw = [];
              setGenerateEnabled();
              resolve();
              return;
            }

            if (missing.length) {
              setStatus("bad",
                `CSV is missing required columns:\n- ${missing.join("\n- ")}\n\nFound columns:\n- ${headers.join("\n- ")}`
              );
            } else {
              setStatus("ok", `CSV loaded.\nRows: ${data.length}\nColumns: ${headers.length}`);
            }

            setGenerateEnabled();
            resolve();
          },
          error: (err) => {
            finished = true;
            clearTimeout(watchdog);
            reject(new Error(err?.message || String(err)));
          }
        });
      });

    } catch (e) {
      if (watchdog) clearTimeout(watchdog);
      setStatus("bad", `CSV parse failed:\n${e?.message || e}`);
      state.csvRowsRaw = [];
      setGenerateEnabled();
    }
  }

  // ---------- Templates ----------
  async function readAsArrayBuffer(file) {
    return await file.arrayBuffer();
  }

  async function handleBuyerTpl(files) {
    const f = files[0];
    if (!f) return;
    state.buyerTplFile = f;
    buyerTplMeta.textContent = `${f.name} (${Math.round(f.size/1024)} KB)`;
    setStatus("warn", "Loading buyer template…");
    try {
      state.buyerTplBuf = await readAsArrayBuffer(f);
      setStatus("ok", "Buyer template loaded.");
    } catch (e) {
      state.buyerTplBuf = null;
      setStatus("bad", `Buyer template load failed:\n${e?.message || e}`);
    }
    setGenerateEnabled();
  }

  async function handleSellerTpl(files) {
    const f = files[0];
    if (!f) return;
    state.sellerTplFile = f;
    sellerTplMeta.textContent = `${f.name} (${Math.round(f.size/1024)} KB)`;
    setStatus("warn", "Loading seller template…");
    try {
      state.sellerTplBuf = await readAsArrayBuffer(f);
      setStatus("ok", "Seller template loaded.");
    } catch (e) {
      state.sellerTplBuf = null;
      setStatus("bad", `Seller template load failed:\n${e?.message || e}`);
    }
    setGenerateEnabled();
  }

  // ---------- DOCX data prep ----------
  function normalizeData(row) {
    const data = {};

    for (const [key, col] of Object.entries(CONFIG.MAP)) {
      let val = getCell(row, col);

      if (!val && CONFIG.FALLBACKS[key]) {
        for (const alt of CONFIG.FALLBACKS[key]) {
          val = getCell(row, alt);
          if (val) break;
        }
      }

      if (key === "price_cwt" || key === "down_money_due") val = moneyClean(val);
      data[key] = val;
    }

    // Compatibility: support both {location} and {Location}
    data.Location = data.location;

    return data;
  }

  // ---------- DOCX generation ----------
  function assertLibsForDocx() {
    if (typeof PizZip === "undefined") {
      throw new Error("PizZip is not loaded. Check /lib/pizzip.min.js path/case.");
    }
    if (typeof window.docxtemplater === "undefined") {
      throw new Error("docxtemplater is not loaded. Check /lib/docxtemplater.min.js path/case.");
    }
  }

  function renderDocx(templateArrayBuffer, data) {
    assertLibsForDocx();

    const zip = new PizZip(templateArrayBuffer);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true
    });

    doc.setData(data);

    try {
      doc.render();
    } catch (error) {
      const e = error;
      const explanation =
        (e.properties && e.properties.errors)
          ? e.properties.errors.map(er => er.properties?.explanation).filter(Boolean).join("\n")
          : (e.message || String(e));

      throw new Error(`Template render failed:\n${explanation}`);
    }

    return doc.getZip().generate({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
  }

  // ---------- Build lots ----------
  function buildLotsFromCsv() {
    const lots = [];
    const rows = state.csvRowsRaw;

    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];

      const contractNo = getCell(r, "Contract #");
      const lotNo = getCell(r, "Lot Number #2") || getCell(r, "Lot Number");
      const buyerName = getCell(r, "Buyer");
      const consignorName = getCell(r, "Consignor");

      // Skip totally empty lines
      if (!contractNo && !lotNo && !buyerName && !consignorName) continue;

      const sellerFilename = sanitizeFilename(`${consignorName || "Consignor"}-${contractNo || "Contract"}.docx`);
      const buyerFilename = sanitizeFilename(`${contractNo || "Contract"}-${buyerName || "Buyer"}.docx`);

      lots.push({
        id: `${contractNo || "NA"}__${lotNo || i}__${i}`,
        idx: i + 1,
        contractNo,
        lotNo,
        buyerName,
        consignorName,
        selected: false,
        buyerDoc: null,
        sellerDoc: null,
        buyerFilename,
        sellerFilename,
        data: normalizeData(r)
      });
    }

    return lots;
  }

  // ---------- Generate ----------
  async function generateAll() {
    if (!canGenerate()) return;

    // enforce required columns exist
    const headers = Object.keys(state.csvRowsRaw[0] || {});
    const missing = requiredMissing(headers);
    if (missing.length) {
      setStatus("bad", `Cannot generate because CSV is missing required columns:\n- ${missing.join("\n- ")}`);
      return;
    }

    setStatus("warn", "Generating contracts…");

    const lots = buildLotsFromCsv();
    if (!lots.length) {
      setStatus("bad", "No usable rows found to generate.");
      state.lots = [];
      renderLots([]);
      setDownloadEnabled(false);
      return;
    }

    let ok = 0;
    const errors = [];

    for (let i = 0; i < lots.length; i++) {
      const lot = lots[i];
      try {
        lot.buyerDoc = renderDocx(state.buyerTplBuf, lot.data);
        lot.sellerDoc = renderDocx(state.sellerTplBuf, lot.data);
        ok++;
      } catch (e) {
        errors.push(`Row ${lot.idx} (Contract # ${lot.contractNo || "?"}, Lot ${lot.lotNo || "?"}): ${e.message || e}`);
      }
    }

    state.lots = lots;
    renderLots(lots);
    setDownloadEnabled(ok > 0);

    if (errors.length) {
      setStatus(
        "warn",
        `Generated ${ok} / ${lots.length} lots.\n\nErrors:\n- ${errors.slice(0, 12).join("\n- ")}${errors.length > 12 ? `\n…plus ${errors.length - 12} more` : ""}`
      );
    } else {
      setStatus("ok", `Generated ${ok} / ${lots.length} lots.\nYou can now download lot-by-lot or ZIP.`);
    }
  }

  // ---------- Results UI ----------
  function getFilteredLots() {
    let lots = [...state.lots];
    if (onlyWithBuyer.checked) lots = lots.filter(l => (l.buyerName || "").trim().length > 0);
    if (onlyWithConsignor.checked) lots = lots.filter(l => (l.consignorName || "").trim().length > 0);
    return lots;
  }

  function syncSelectAllCheckbox() {
    const visible = getFilteredLots();
    if (!visible.length) {
      selectAll.checked = false;
      selectAll.indeterminate = false;
      return;
    }
    const selectedCount = visible.filter(l => l.selected).length;
    selectAll.checked = selectedCount === visible.length;
    selectAll.indeterminate = selectedCount > 0 && selectedCount < visible.length;
  }

  function renderLots(lotsInput) {
    const lots = lotsInput || [];
    lotCount.textContent = String(lots.length);

    if (!lots.length) {
      resultsBox.innerHTML = `<div class="lotrow"><div class="small">No generated lots yet.</div></div>`;
      return;
    }

    const lotsToShow = getFilteredLots();

    const html = lotsToShow.map(lot => {
      const buyerOk = !!lot.buyerDoc;
      const sellerOk = !!lot.sellerDoc;

      const left = `
        <div class="c1">
          <input type="checkbox" data-id="${lot.id}" class="lotSelect" ${lot.selected ? "checked" : ""} />
        </div>
      `;

      const c2 = `
        <div class="c2">
          <div><b class="mono">${escapeHtml(lot.contractNo || "—")}</b> <span class="small">• Lot</span> <b class="mono">${escapeHtml(lot.lotNo || "—")}</b></div>
          <div class="small">Row ${lot.idx}</div>
        </div>
      `;

      const c3 = `
        <div class="c3">
          <div class="small">Buyer</div>
          <div>${escapeHtml(lot.buyerName || "—")}</div>
        </div>
      `;

      const c4 = `
        <div class="c4">
          <div class="small">Consignor</div>
          <div>${escapeHtml(lot.consignorName || "—")}</div>
        </div>
      `;

      const c5 = `
        <div class="c5 actions">
          <button class="btn" data-action="dlBuyer" data-id="${lot.id}" ${buyerOk ? "" : "disabled"}>Download Buyer</button>
          <button class="btn" data-action="dlSeller" data-id="${lot.id}" ${sellerOk ? "" : "disabled"}>Download Seller</button>
        </div>
      `;

      return `<div class="lotrow">${left}${c2}${c3}${c4}${c5}</div>`;
    }).join("");

    resultsBox.innerHTML = html;

    resultsBox.querySelectorAll(".lotSelect").forEach(cb => {
      cb.addEventListener("change", (e) => {
        const id = e.target.getAttribute("data-id");
        const lot = state.lots.find(x => x.id === id);
        if (lot) lot.selected = e.target.checked;
        syncSelectAllCheckbox();
      });
    });

    resultsBox.querySelectorAll("button[data-action]").forEach(btn => {
      btn.addEventListener("click", (e) => {
        const action = e.target.getAttribute("data-action");
        const id = e.target.getAttribute("data-id");
        const lot = state.lots.find(x => x.id === id);
        if (!lot) return;

        if (action === "dlBuyer") downloadOne(lot, "buyer");
        if (action === "dlSeller") downloadOne(lot, "seller");
      });
    });

    syncSelectAllCheckbox();
  }

  selectAll.addEventListener("change", () => {
    const visible = getFilteredLots();
    visible.forEach(l => l.selected = selectAll.checked);
    renderLots(state.lots);
  });

  onlyWithBuyer.addEventListener("change", () => renderLots(state.lots));
  onlyWithConsignor.addEventListener("change", () => renderLots(state.lots));

  // ---------- Downloads ----------
  function selectedOrAllLots() {
    const selected = state.lots.filter(l => l.selected);
    return selected.length ? selected : state.lots;
  }

  function downloadOne(lot, which) {
    if (which === "buyer") {
      if (!lot.buyerDoc) return;
      saveAs(lot.buyerDoc, lot.buyerFilename);
    } else if (which === "seller") {
      if (!lot.sellerDoc) return;
      saveAs(lot.sellerDoc, lot.sellerFilename);
    }
  }

  async function downloadZip(mode) {
    // mode: "buyer" | "seller" | "all"
    if (typeof JSZip === "undefined") {
      setStatus("bad", "JSZip is not loaded. Check /lib/jszip.min.js path/case.");
      return;
    }
    if (typeof saveAs === "undefined") {
      setStatus("bad", "FileSaver (saveAs) is not loaded. Check /lib/FileSaver.min.js path/case.");
      return;
    }

    const lots = selectedOrAllLots().filter(l => {
      if (mode === "buyer") return !!l.buyerDoc;
      if (mode === "seller") return !!l.sellerDoc;
      return !!l.buyerDoc || !!l.sellerDoc;
    });

    if (!lots.length) {
      setStatus("bad", "No contracts available for that download mode (check your selection/filters).");
      return;
    }

    setStatus("warn", "Building ZIP…");

    const zip = new JSZip();
    const stamp = new Date();
    const yyyy = stamp.getFullYear();
    const mm = String(stamp.getMonth() + 1).padStart(2, "0");
    const dd = String(stamp.getDate()).padStart(2, "0");
    const dateTag = `${yyyy}-${mm}-${dd}`;

    const buyerFolder = zip.folder("Buyer Contracts");
    const sellerFolder = zip.folder("Seller Contracts");

    for (const lot of lots) {
      if ((mode === "buyer" || mode === "all") && lot.buyerDoc) {
        buyerFolder.file(lot.buyerFilename, lot.buyerDoc);
      }
      if ((mode === "seller" || mode === "all") && lot.sellerDoc) {
        sellerFolder.file(lot.sellerFilename, lot.sellerDoc);
      }
    }

    if (mode === "buyer") zip.remove("Seller Contracts");
    if (mode === "seller") zip.remove("Buyer Contracts");

    const blob = await zip.generateAsync({ type: "blob" });

    const zipName =
      mode === "buyer" ? `Buyer_Contracts_${dateTag}.zip` :
      mode === "seller" ? `Seller_Contracts_${dateTag}.zip` :
      `All_Contracts_${dateTag}.zip`;

    saveAs(blob, zipName);
    setStatus("ok", `ZIP downloaded: ${zipName}`);
  }

  zipBuyerBtn.addEventListener("click", () => downloadZip("buyer"));
  zipSellerBtn.addEventListener("click", () => downloadZip("seller"));
  zipAllBtn.addEventListener("click", () => downloadZip("all"));

  // ---------- Clear ----------
  function clearSessionFiles(silent=false) {
    state.csvFile = null;
    state.csvRowsRaw = [];

    state.buyerTplFile = null;
    state.sellerTplFile = null;
    state.buyerTplBuf = null;
    state.sellerTplBuf = null;

    state.lots = [];

    csvInput.value = "";
    buyerTplInput.value = "";
    sellerTplInput.value = "";

    csvMeta.textContent = "No file selected";
    buyerTplMeta.textContent = "No template selected";
    sellerTplMeta.textContent = "No template selected";

    setGenerateEnabled();
    setDownloadEnabled(false);
    renderLots([]);

    if (!silent) setStatus("ok", "Session cleared. Re-upload CSV and templates to generate again.");
  }

  clearBtn.addEventListener("click", () => clearSessionFiles());

  // ---------- Wire dropzones ----------
  wireDropzone(dzCsv, csvInput, handleCsv);
  wireDropzone(dzBuyerTpl, buyerTplInput, handleBuyerTpl);
  wireDropzone(dzSellerTpl, sellerTplInput, handleSellerTpl);

  // ---------- Generate button ----------
  generateBtn.addEventListener("click", generateAll);

  // ---------- Init ----------
  function init() {
    renderLots([]);
    setDownloadEnabled(false);
    setGenerateEnabled();

    if (isUnlocked()) unlock();
    else lock();

    // Extra: show if libs are missing immediately (helps catch 404)
    const missing = [];
    if (typeof Papa === "undefined") missing.push("PapaParse (Papa) missing");
    if (typeof PizZip === "undefined") missing.push("PizZip missing");
    if (typeof window.docxtemplater === "undefined") missing.push("docxtemplater missing");
    if (typeof JSZip === "undefined") missing.push("JSZip missing");
    if (typeof saveAs === "undefined") missing.push("FileSaver (saveAs) missing");

    if (missing.length) {
      setStatus(
        "warn",
        "One or more libraries did not load:\n- " + missing.join("\n- ") +
        "\n\nOpen DevTools → Network and confirm /lib/*.js files return 200 (not 404)."
      );
    }
  }

  function wireDropzone(zoneEl, inputEl, onFiles) {
    zoneEl.addEventListener("dragover", (e) => {
      e.preventDefault();
      zoneEl.style.borderColor = "rgba(51,102,153,.75)";
    });
    zoneEl.addEventListener("dragleave", () => {
      zoneEl.style.borderColor = "rgba(51,102,153,.35)";
    });
    zoneEl.addEventListener("drop", (e) => {
      e.preventDefault();
      zoneEl.style.borderColor = "rgba(51,102,153,.35)";
      const files = e.dataTransfer?.files;
      if (files && files.length) onFiles(files);
    });
    inputEl.addEventListener("change", () => {
      const files = inputEl.files;
      if (files && files.length) onFiles(files);
    });
  }

  init();
})();
