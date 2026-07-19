const tabs = document.querySelectorAll(".tab");
const panels = document.querySelectorAll(".tab-panel");
const search = document.getElementById("tableSearch");
const pdfInput = document.getElementById("pdfInput");
const selectedFiles = document.getElementById("selectedFiles");
const analyzeButton = document.getElementById("analyzeButton");
const uploadForm = document.getElementById("uploadForm");
const progressPanel = document.getElementById("progressPanel");
const metricButtons = document.querySelectorAll(".metric-button");
let statusFilter = null;

function activePanel() {
  return document.querySelector(".tab-panel.active");
}

function filterRows() {
  if (!search) return;
  const query = search.value.trim().toLowerCase();
  const panel = activePanel();
  if (!panel) return;
  panel.querySelectorAll("tbody tr").forEach((row) => {
    const text = row.textContent.toLowerCase();
    const matchesSearch = !query || text.includes(query);
    const matchesStatus = !statusFilter || statusFilter.some((term) => text.includes(term.toLowerCase()));
    row.style.display = matchesSearch && matchesStatus ? "" : "none";
  });
}

function activateTab(tabName) {
  const tab = document.querySelector(`.tab[data-tab="${tabName}"]`);
  const panel = document.getElementById(tabName);
  if (!tab || !panel) return;
  tabs.forEach((item) => item.classList.remove("active"));
  panels.forEach((item) => item.classList.remove("active"));
  tab.classList.add("active");
  panel.classList.add("active");
}

function applyStatusFilter(type) {
  const badTerms = ["DIFFERENT", "NON CONFORME"];
  const verifyTerms = ["A VERIFIER"];
  const terms = type === "bad" ? badTerms : verifyTerms;

  const coverHasMatch = Array.from(document.querySelectorAll("#cover tbody tr"))
    .some((row) => terms.some((term) => row.textContent.toUpperCase().includes(term)));
  const inspectionHasMatch = Array.from(document.querySelectorAll("#inspection tbody tr"))
    .some((row) => terms.some((term) => row.textContent.toUpperCase().includes(term)));
  const standardsHasMatch = Array.from(document.querySelectorAll("#standards tbody tr"))
    .some((row) => type === "verify" && row.textContent.toUpperCase().includes("NA"));

  if (coverHasMatch) {
    activateTab("cover");
  } else if (inspectionHasMatch) {
    activateTab("inspection");
  } else if (standardsHasMatch) {
    activateTab("standards");
  }

  statusFilter = terms;
  if (search) search.value = "";
  filterRows();
  const results = document.querySelector(".results");
  if (results) results.scrollIntoView({ behavior: "smooth", block: "start" });
}

tabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    tabs.forEach((item) => item.classList.remove("active"));
    panels.forEach((panel) => panel.classList.remove("active"));
    tab.classList.add("active");
    document.getElementById(tab.dataset.tab).classList.add("active");
    if (search) search.value = "";
    statusFilter = null;
    filterRows();
  });
});

if (search) {
  search.addEventListener("input", () => {
    statusFilter = null;
    filterRows();
  });
}

metricButtons.forEach((button) => {
  button.addEventListener("click", () => {
    applyStatusFilter(button.dataset.filter);
  });
});

if (pdfInput && selectedFiles && analyzeButton) {
  pdfInput.addEventListener("change", () => {
    const files = Array.from(pdfInput.files || []);
    const pdfs = files.filter((file) => file.name.toLowerCase().endsWith(".pdf"));
    analyzeButton.disabled = pdfs.length === 0;
    selectedFiles.textContent = pdfs.length
      ? pdfs.map((file) => file.name).join(" | ")
      : "Aucun fichier sélectionné";
  });
}

if (uploadForm && analyzeButton) {
  uploadForm.addEventListener("submit", () => {
    analyzeButton.disabled = true;
    analyzeButton.textContent = "Analyse en cours...";
    if (progressPanel) progressPanel.hidden = false;
  });
}

document.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") return;
  if (!uploadForm || !analyzeButton || analyzeButton.disabled) return;
  const active = document.activeElement;
  if (active && ["TEXTAREA", "BUTTON", "A"].includes(active.tagName)) return;
  event.preventDefault();
  uploadForm.requestSubmit();
});
