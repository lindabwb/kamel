const tabs = document.querySelectorAll(".tab");
const panels = document.querySelectorAll(".tab-panel");
const search = document.getElementById("tableSearch");
const clearFilterBtn = document.getElementById("clearFilter");
const pdfInput = document.getElementById("pdfInput");
const selectedFiles = document.getElementById("selectedFiles");
const analyzeButton = document.getElementById("analyzeButton");
const uploadForm = document.getElementById("uploadForm");
const progressPanel = document.getElementById("progressPanel");
const metricButtons = document.querySelectorAll(".metric-clickable");
let currentFilter = null; // 'all', 'conforme', 'non-conforme', 'a-verifier'

function activePanel() {
  return document.querySelector(".tab-panel.active");
}

function filterRows() {
  const query = search ? search.value.trim().toLowerCase() : "";
  const panel = activePanel();
  if (!panel) return;
  
  panel.querySelectorAll("tbody tr").forEach((row) => {
    const text = row.textContent.toLowerCase();
    const matchesSearch = !query || text.includes(query);
    let matchesStatus = true;
    
    if (currentFilter && currentFilter !== 'all') {
      const rowStatus = row.dataset.status || '';
      if (currentFilter === 'conforme') {
        matchesStatus = rowStatus === 'conforme' || rowStatus === 'ok';
      } else if (currentFilter === 'non-conforme') {
        matchesStatus = rowStatus === 'non-conforme' || rowStatus === 'different';
      } else if (currentFilter === 'a-verifier') {
        matchesStatus = rowStatus === 'a-verifier';
      }
    }
    
    row.style.display = matchesSearch && matchesStatus ? "" : "none";
  });
  
  // Mettre à jour les métriques pour montrer le filtre actif
  metricButtons.forEach(btn => {
    btn.classList.remove('active');
    if (btn.dataset.filter === currentFilter) {
      btn.classList.add('active');
    }
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
  // Re-appliquer le filtre sur le nouveau panel
  filterRows();
}

function applyFilter(filterType) {
  currentFilter = filterType;
  if (search) search.value = "";
  filterRows();
  // Scroll vers les résultats
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
    filterRows();
  });
});

if (search) {
  search.addEventListener("input", () => {
    filterRows();
  });
}

if (clearFilterBtn) {
  clearFilterBtn.addEventListener("click", () => {
    currentFilter = null;
    if (search) search.value = "";
    filterRows();
    metricButtons.forEach(btn => btn.classList.remove('active'));
  });
}

metricButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const filter = button.dataset.filter;
    // Si on clique sur le même filtre, on le désactive
    if (currentFilter === filter) {
      currentFilter = null;
      metricButtons.forEach(btn => btn.classList.remove('active'));
    } else {
      applyFilter(filter);
    }
    filterRows();
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