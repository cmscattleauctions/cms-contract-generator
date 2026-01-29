/* ============================================================
   CMS Contract Generator — FULL app.js
   ------------------------------------------------------------
   - Client-only (GitHub Pages safe)
   - Requires:
       CSV upload
       Buyer DOCX template upload
       Seller DOCX template upload
   - Generates:
       Buyer + Seller contracts per lot
   - Downloads:
       • Per-lot Buyer
       • Per-lot Seller
       • Buyer ZIP
       • Seller ZIP
       • All ZIP
   - PIN gate: 0623
   - PapaParse is assumed loaded BEFORE this file
   ============================================================ */

(() => {
  const CONFIG = {
    PIN: "0623",

    REQUIRED_COLS: [
      "Contract #",
      "Consignor",
      "Buyer",
      "Lot Number #2"
    ],

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
    }
  };

  /* ================= DOM ================= */
  const $ = id => document.getElementById(id);

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

  /* ================= STATE ================= */
  const state = {
    csvRows: [],
    buyerTplBuf: null,
    sellerTplBuf: null,
    lots: []
  };

  /* ================= HELPERS ================= */
  function setStatus(type, msg) {
    validationBox.style.display = "block";
    validationBox.className = `status ${type}`;
    validationBox.textContent = msg;
  }

  function sanitizeFilename(name) {
    return String(name || "Untitled")
      .replace(/[\/\\:*?"<>|]+/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  }

  function escapeHtml(s) {
    return String(s || "")
      .replaceAll("&","&amp;")
      .replaceAll("<","&lt;")
      .replaceAll(">","&gt;")
      .replaceAll('"',"&quot;");
  }

  function getCell(row, col) {
    return row[col] == null ? "" : String(row[col]).trim();
  }

  function cleanMoney(v) {
    return String(v || "").replace(/[$,]/g, "").trim();
  }

  function canGenerate() {
    return state.csvRows.length && state.buyerTplBuf && state.sellerTplBuf;
  }

  /* ================= PIN ================= */
  function lock() {
    sessionStorage.removeItem("cms_pin_ok");
    pinOverlay.style.display = "flex";
    sessionPill.textContent = "Locked";
  }

  function unlock() {
    sessionStorage.setItem("cms_pin_ok","1");
    pinOverlay.style.display = "none";
    sessionPill.textContent = "Unlocked";
  }

  pinUnlockBtn.onclick = () => {
    if (pinInput.value === CONFIG.PIN) unlock();
    else pinErr.style.display = "block";
  };

  pinClearBtn.onclick = () => {
    pinInput.value = "";
    pinErr.style.display = "none";
  };

  exitBtn.onclick = () => {
    clearAll(true);
    lock();
  };

  /* ================= FILE LOADERS ================= */
  function readBuf(file) {
    return file.arrayBuffer();
  }

  csvInput.onchange = () => parseCsv(csvInput.files[0]);
  buyerTplInput.onchange = async () => {
    state.buyerTplBuf = await readBuf(buyerTplInput.files[0]);
    buyerTplMeta.textContent = buyerTplInput.files[0].name;
    setStatus("ok","Buyer template loaded.");
    generateBtn.disabled = !canGenerate();
  };
  sellerTplInput.onchange = async () => {
    state.sellerTplBuf = await readBuf(sellerTplInput.files[0]);
    sellerTplMeta.textContent = sellerTplInput.files[0].name;
    setStatus("ok","Seller template loaded.");
    generateBtn.disabled = !canGenerate();
  };

  function parseCsv(file) {
    setStatus("warn","Parsing CSV…");

    Papa.parse(file,{
      header:true,
      skipEmptyLines:true,
      worker:true,
      complete:(res)=>{
        state.csvRows = res.data || [];
        csvMeta.textContent = file.name;

        const headers = res.meta.fields || [];
        const missing = CONFIG.REQUIRED_COLS.filter(c=>!headers.includes(c));

        if (missing.length) {
          setStatus("bad","Missing required columns:\n"+missing.join("\n"));
          state.csvRows=[];
          return;
        }

        setStatus("ok",`CSV loaded. Rows: ${state.csvRows.length}`);
        generateBtn.disabled = !canGenerate();
      },
      error:(err)=>{
        setStatus("bad","CSV parse error:\n"+err.message);
      }
    });
  }

  /* ================= DOCX ================= */
  function renderDocx(buf,data) {
    const zip = new PizZip(buf);
    const doc = new window.docxtemplater(zip,{paragraphLoop:true,linebreaks:true});
    doc.setData(data);
    doc.render();
    return doc.getZip().generate({
      type:"blob",
      mimeType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    });
  }

  function normalize(row) {
    const d = {};
    for (const [k,c] of Object.entries(CONFIG.MAP)) {
      let v = getCell(row,c);
      if (k === "price_cwt" || k === "down_money_due") v = cleanMoney(v);
      d[k] = v;
    }
    d.Location = d.location;
    return d;
  }

  /* ================= GENERATE ================= */
  generateBtn.onclick = async () => {
    setStatus("warn","Generating contracts…");
    state.lots = [];

    for (let i=0;i<state.csvRows.length;i++) {
      const r = state.csvRows[i];
      const data = normalize(r);

      const contractNo = data.contract_no;
      const buyer = data.buyer;
      const consignor = data.consignor;

      const buyerName = sanitizeFilename(`${contractNo}-${buyer}.docx`);
      const sellerName = sanitizeFilename(`${consignor}-${contractNo}.docx`);

      try {
        const buyerDoc = renderDocx(state.buyerTplBuf,data);
        const sellerDoc = renderDocx(state.sellerTplBuf,data);

        state.lots.push({
          id:i,
          contractNo,
          buyer,
          consignor,
          buyerDoc,
          sellerDoc,
          buyerName,
          sellerName,
          selected:false
        });
      } catch (e) {
        console.error(e);
      }
    }

    renderLots();
    zipBuyerBtn.disabled = zipSellerBtn.disabled = zipAllBtn.disabled = !state.lots.length;
    setStatus("ok",`Generated ${state.lots.length} lots.`);
  };

  /* ================= UI ================= */
  function renderLots() {
    lotCount.textContent = state.lots.length;
    resultsBox.innerHTML = "";

    state.lots.forEach(l=>{
      const row = document.createElement("div");
      row.className = "lotrow";
      row.innerHTML = `
        <input type="checkbox" data-id="${l.id}">
        <div><b>${escapeHtml(l.contractNo)}</b></div>
        <div>${escapeHtml(l.buyer)}</div>
        <div>${escapeHtml(l.consignor)}</div>
        <div class="actions">
          <button data-b="${l.id}">Buyer</button>
          <button data-s="${l.id}">Seller</button>
        </div>
      `;
      resultsBox.appendChild(row);

      row.querySelector("[data-b]").onclick=()=>saveAs(l.buyerDoc,l.buyerName);
      row.querySelector("[data-s]").onclick=()=>saveAs(l.sellerDoc,l.sellerName);
      row.querySelector("input").onchange=e=>l.selected=e.target.checked;
    });
  }

  /* ================= ZIP ================= */
  function selectedLots() {
    const s = state.lots.filter(l=>l.selected);
    return s.length ? s : state.lots;
  }

  async function zip(mode) {
    const zip = new JSZip();
    const b = zip.folder("Buyer Contracts");
    const s = zip.folder("Seller Contracts");

    selectedLots().forEach(l=>{
      if ((mode==="buyer"||mode==="all") && l.buyerDoc) b.file(l.buyerName,l.buyerDoc);
      if ((mode==="seller"||mode==="all") && l.sellerDoc) s.file(l.sellerName,l.sellerDoc);
    });

    if (mode==="buyer") zip.remove("Seller Contracts");
    if (mode==="seller") zip.remove("Buyer Contracts");

    const blob = await zip.generateAsync({type:"blob"});
    saveAs(blob,`${mode}_contracts.zip`);
  }

  zipBuyerBtn.onclick = ()=>zip("buyer");
  zipSellerBtn.onclick = ()=>zip("seller");
  zipAllBtn.onclick = ()=>zip("all");

  /* ================= CLEAR ================= */
  function clearAll(silent=false) {
    state.csvRows=[];
    state.buyerTplBuf=null;
    state.sellerTplBuf=null;
    state.lots=[];
    csvInput.value="";
    buyerTplInput.value="";
    sellerTplInput.value="";
    resultsBox.innerHTML="";
    generateBtn.disabled=true;
    zipBuyerBtn.disabled=zipSellerBtn.disabled=zipAllBtn.disabled=true;
    if(!silent) setStatus("ok","Cleared.");
  }

  clearBtn.onclick = ()=>clearAll();

  /* ================= INIT ================= */
  if (sessionStorage.getItem("cms_pin_ok")==="1") unlock();
  else lock();
})();
