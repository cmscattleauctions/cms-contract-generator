/* CMS Contract Generator — client-only
   - PIN gate: 0623
   - Upload CSV + Buyer DOCX template + Seller DOCX template each session
   - For each CSV row: generate Buyer + Seller editable DOCX using docxtemplater with {{ }} delimiters
   - Zip output:
     /Buyer Contracts/{Contract#}-{Buyer}.docx
     /Seller Contracts/{Consignor}-{Contract#}.docx
*/

(() => {
  const CONFIG = {
    PIN: "0623",
    ZIP_FOLDERS: {
      buyer: "Buyer Contracts",
      seller: "Seller Contracts",
    },
    // Required CSV headers (must exist in the CSV file)
    REQUIRED_COLS: [
      "Contract #",
      "Consignor",
      "Buyer",
      "Lot Number #2",
      "Head Count",
    ],
    // docxtemplater delimiters to match your templates' {{Tag}} style
    DELIMS: { start: "{{", end: "}}" },
  };

  // ---------- DOM ----------
  const pinScreen = document.getElementById("pinScreen");
  const appScreen = document.getElementById("appScreen");
  const pinInput  = document.getElementById("pinInput");
  const pinBtn    = document.getElementById("pinBtn");
  const pinMsg    = document.getElementById("pinMsg");

  const dropCsv    = document.getElementById("dropCsv");
  const dropBuyer  = document.getElementById("dropBuyer");
  const dropSeller = document.getElementById("dropSeller");

  const csvPicker = document.getElementById("csvPicker");
  const buyerPicker = document.getElementById("buyerPicker");
  const sellerPicker = document.getElementById("sellerPicker");

  const csvPickBtn = document.getElementById("csvPickBtn");
  const buyerPickBtn = document.getElementById("buyerPickBtn");
  const sellerPickBtn = document.getElementById("sellerPickBtn");

  const csvPill   = document.getElementById("csvPill");
  const buyerPill = document.getElementById("buyerPill");
  const sellerPill= document.getElementById("sellerPill");

  const genBtn   = document.getElementById("genBtn");
  const resetBtn = document.getElementById("resetBtn");

  const statusEl = document.getElementById("status");

  // ---------- STATE ----------
  let csvRows = [];
  let csvFileName = "";

  // We store templates in memory only (not hosted)
  let buyerTemplateBytes = null;  // Uint8Array
  let sellerTemplateBytes = null; // Uint8Array
  let buyerTemplateName = "";
  let sellerTemplateName = "";

  // ---------- UTIL ----------
  function log(msg) {
    statusEl.textContent = `${msg}\n\n` + statusEl.textContent;
  }
  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function setPill(pillEl, type, text) {
    pillEl.classList.remove("ok","bad","warn");
    pillEl.classList.add(type);
    pillEl.textContent = text;
  }

  function sanitizeFilePart(s) {
    const str = String(s ?? "").trim();
    if (!str) return "UNKNOWN";
    // Windows invalid filename chars + trim dots/spaces
    return str
      .replace(/[\/\\:*?"<>|]/g, "-")
      .replace(/\s+/g, " ")
      .trim()
      .replace(/^[.\s]+|[.\s]+$/g, "");
  }

  function canGenerate() {
    return csvRows.length > 0 && buyerTemplateBytes && sellerTemplateBytes;
  }
  function refreshGenerateButton() {
    genBtn.disabled = !canGenerate();
  }

  function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
  }
  function addDragUI(el, on) {
    el.classList.toggle("dragover", !!on);
  }

  function fileToUint8Array(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => resolve(new Uint8Array(r.result));
      r.onerror = () => reject(new Error("Failed reading file."));
      r.readAsArrayBuffer(file);
    });
  }

  // ---------- PIN ----------
  function unlock() {
    pinScreen.classList.add("hide");
    appScreen.classList.remove("hide");
    setStatus("Unlocked. Upload CSV + Buyer template + Seller template to enable ZIP generation.");
  }

  pinBtn.addEventListener("click", () => {
    const val = String(pinInput.value || "").trim();
    if (val === CONFIG.PIN) {
      pinMsg.textContent = "";
      unlock();
    } else {
      pinMsg.textContent = "Incorrect PIN.";
      pinMsg.style.color = "#fecaca";
    }
  });

  pinInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter") pinBtn.click();
  });

  // ---------- CSV PARSE ----------
  async function handleCsvFile(file) {
    setStatus("Reading CSV...");
    csvFileName = file.name;

    const text = await file.text();

    const parsed = Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false, // keep values as strings as much as possible
    });

    if (parsed.errors?.length) {
      setPill(csvPill, "bad", "CSV parse error");
      setStatus("CSV parse errors:\n" + parsed.errors.map(e => `${e.message} (row ${e.row})`).join("\n"));
      csvRows = [];
      refreshGenerateButton();
      return;
    }

    const rows = (parsed.data || []).filter(r => {
      // filter out completely blank rows
      return Object.values(r).some(v => String(v ?? "").trim() !== "");
    });

    // Validate required columns exist
    const headers = parsed.meta?.fields || [];
    const missing = CONFIG.REQUIRED_COLS.filter(c => !headers.includes(c));

    if (missing.length) {
      setPill(csvPill, "bad", "Missing columns");
      setStatus(
        "CSV is missing required columns:\n" +
        missing.map(m => `- ${m}`).join("\n") +
        "\n\nFix the CSV headers (or tell me and I’ll update mapping)."
      );
      csvRows = [];
      refreshGenerateButton();
      return;
    }

    csvRows = rows;
    setPill(csvPill, "ok", `${file.name} (${rows.length} rows)`);
    setStatus(`CSV loaded: ${rows.length} rows.\nNow upload BOTH DOCX templates.`);
    refreshGenerateButton();
  }

  // ---------- TEMPLATE UPLOAD ----------
  async function handleBuyerTemplate(file) {
    setStatus("Reading Buyer template...");
    buyerTemplateBytes = await fileToUint8Array(file);
    buyerTemplateName = file.name;
    setPill(buyerPill, "ok", `${file.name}`);
    log("Buyer template loaded.");
    refreshGenerateButton();
  }

  async function handleSellerTemplate(file) {
    setStatus("Reading Seller template...");
    sellerTemplateBytes = await fileToUint8Array(file);
    sellerTemplateName = file.name;
    setPill(sellerPill, "ok", `${file.name}`);
    log("Seller template loaded.");
    refreshGenerateButton();
  }

  // ---------- DRAG/DROP WIRING ----------
  function wireDropZone(el, acceptFn) {
    ["dragenter","dragover","dragleave","drop"].forEach(evt => {
      el.addEventListener(evt, preventDefaults, false);
    });
    ["dragenter","dragover"].forEach(evt => {
      el.addEventListener(evt, () => addDragUI(el, true), false);
    });
    ["dragleave","drop"].forEach(evt => {
      el.addEventListener(evt, () => addDragUI(el, false), false);
    });
    el.addEventListener("drop", async (e) => {
      const file = e.dataTransfer.files?.[0];
      if (!file) return;
      await acceptFn(file);
    });
  }

  wireDropZone(dropCsv, async (file) => {
    if (!file.name.toLowerCase().endsWith(".csv")) {
      setStatus("That file is not a CSV.");
      return;
    }
    await handleCsvFile(file);
  });

  wireDropZone(dropBuyer, async (file) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setStatus("Buyer template must be a .docx file.");
      return;
    }
    await handleBuyerTemplate(file);
  });

  wireDropZone(dropSeller, async (file) => {
    if (!file.name.toLowerCase().endsWith(".docx")) {
      setStatus("Seller template must be a .docx file.");
      return;
    }
    await handleSellerTemplate(file);
  });

  // File pick buttons
  csvPickBtn.addEventListener("click", () => csvPicker.click());
  buyerPickBtn.addEventListener("click", () => buyerPicker.click());
  sellerPickBtn.addEventListener("click", () => sellerPicker.click());

  csvPicker.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (f) await handleCsvFile(f);
    csvPicker.value = "";
  });
  buyerPicker.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (f) await handleBuyerTemplate(f);
    buyerPicker.value = "";
  });
  sellerPicker.addEventListener("change", async (e) => {
    const f = e.target.files?.[0];
    if (f) await handleSellerTemplate(f);
    sellerPicker.value = "";
  });

  // ---------- DOCX RENDER ----------
  function renderDocxFromTemplate(templateBytes, dataObj) {
    // Clone bytes because docxtemplater mutates zip state
    const bytes = templateBytes.slice(0);

    const zip = new PizZip(bytes);
    const doc = new window.docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: CONFIG.DELIMS, // match {{Tag}}
    });

    doc.render(dataObj);

    return doc.getZip().generate({ type: "blob" });
  }

  // ---------- GENERATE ZIP ----------
  genBtn.addEventListener("click", async () => {
    if (!canGenerate()) {
      setStatus("Upload CSV + Buyer template + Seller template first.");
      return;
    }

    setStatus("Generating DOCX files and building ZIP...");
    genBtn.disabled = true;

    const zip = new JSZip();
    const buyerFolder = zip.folder(CONFIG.ZIP_FOLDERS.buyer);
    const sellerFolder = zip.folder(CONFIG.ZIP_FOLDERS.seller);

    let okCount = 0;
    let failCount = 0;
    const usedNames = new Set();

    for (let i = 0; i < csvRows.length; i++) {
      const row = csvRows[i];

      try {
        // Data object keys MUST match the DOCX tags inside {{ }}
        // Your templates use tags like {{Contract #}}, {{Buyer}}, etc.
        // We also set both Shrink and shrink because seller doc uses lowercase {{shrink}}.
        const data = {
          ...row,
          Shrink: row["Shrink"],
          shrink: row["Shrink"],
        };

        const contractNo = sanitizeFilePart(row["Contract #"]);
        const buyerName  = sanitizeFilePart(row["Buyer"]);
        const consignor  = sanitizeFilePart(row["Consignor"]);

        // File naming rules you specified:
        // Seller: Consignor-ContractNumber
        // Buyer: ContractNumber-Buyer
        let sellerName = `${consignor}-${contractNo}.docx`;
        let buyerNameFile = `${contractNo}-${buyerName}.docx`;

        // Prevent collisions (if same names repeat)
        sellerName = dedupeName(sellerName, usedNames);
        buyerNameFile = dedupeName(buyerNameFile, usedNames);

        const buyerBlob = renderDocxFromTemplate(buyerTemplateBytes, data);
        const sellerBlob = renderDocxFromTemplate(sellerTemplateBytes, data);

        buyerFolder.file(buyerNameFile, buyerBlob);
        sellerFolder.file(sellerName, sellerBlob);

        okCount++;
      } catch (err) {
        failCount++;
        log(`Row ${i + 1} failed: ${err?.message || err}`);
      }
    }

    const zipName = `CMS Contracts - ${new Date().toISOString().slice(0,10)}.zip`;

    try {
      const outBlob = await zip.generateAsync({ type: "blob" });
      saveAs(outBlob, zipName);
      setStatus(
        `Done.\n` +
        `Rows processed: ${csvRows.length}\n` +
        `Contracts generated: ${okCount * 2} (Buyer+Seller per row)\n` +
        `Row failures: ${failCount}\n\n` +
        `Downloaded: ${zipName}`
      );
    } catch (e) {
      setStatus("ZIP build failed: " + (e?.message || e));
    }

    refreshGenerateButton();
  });

  function dedupeName(filename, usedSet) {
    const base = filename.replace(/\.docx$/i, "");
    let name = filename;
    let n = 2;
    while (usedSet.has(name.toLowerCase())) {
      name = `${base} (${n}).docx`;
      n++;
    }
    usedSet.add(name.toLowerCase());
    return name;
  }

  // ---------- RESET ----------
  resetBtn.addEventListener("click", () => {
    csvRows = [];
    csvFileName = "";
    buyerTemplateBytes = null;
    sellerTemplateBytes = null;
    buyerTemplateName = "";
    sellerTemplateName = "";

    setPill(csvPill, "warn", "No CSV loaded");
    setPill(buyerPill, "warn", "No Buyer template loaded");
    setPill(sellerPill, "warn", "No Seller template loaded");

    setStatus("Reset complete. Re-upload files to generate again.");
    refreshGenerateButton();
  });

  // Initial UI state
  setPill(csvPill, "warn", "No CSV loaded");
  setPill(buyerPill, "warn", "No Buyer template loaded");
  setPill(sellerPill, "warn", "No Seller template loaded");
  refreshGenerateButton();

  // Helpful startup log
  console.log("CMS Contract Generator loaded.");
})();

