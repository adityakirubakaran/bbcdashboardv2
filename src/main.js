import { parseExcel, processData, processSalesData } from './dataParser.js';
import { renderCharts } from './chartConfig.js';

window.appDatasets = [];

const uploadInput = document.getElementById('excel-upload');
const noDataPlaceholder = document.getElementById('no-data-placeholder');
const mainChartCanvas = document.getElementById('main-chart');

uploadInput.addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  try {
    const data = await parseExcel(file);
    window.appDatasets.push({
      id: Date.now().toString(),
      name: file.name,
      type: data.type,
      rawRows: data.rawRows,
      rawHeaders: data.rawHeaders,
      uniqueSites: data.uniqueSites || [],
      uniqueCategories: data.uniqueCategories || []
    });
    
    renderDatasetManager();
    updateDashboard();
  } catch (err) {
    console.error("Error parsing the excel file:", err);
    alert("There was an error parsing the uploaded file. Please ensure it is the correct format.");
  }
  
  // Allow re-uploading the same file without forcing the user to pick a different one first.
  e.target.value = '';
});

function renderDatasetManager() {
    const container = document.getElementById('uploaded-sheets');
    container.innerHTML = '';
    
    window.appDatasets.forEach(ds => {
        const pill = document.createElement('div');
        pill.className = 'dataset-pill';
        pill.innerHTML = `
            <span>${ds.name}</span>
            <button class="dataset-pill-remove" aria-label="Remove dataset">×</button>
        `;
        
        pill.querySelector('.dataset-pill-remove').addEventListener('click', () => {
            window.appDatasets = window.appDatasets.filter(d => d.id !== ds.id);
            renderDatasetManager();
            updateDashboard();
        });
        
        container.appendChild(pill);
    });
}

function populateFilters(sites, categories) {
    const locContainer = document.getElementById('location-filters-list');
    const catContainer = document.getElementById('category-filters-list');
    
    if(locContainer) locContainer.innerHTML = '';
    if(catContainer) catContainer.innerHTML = '';

    sites.forEach(site => {
        const item = document.createElement('div');
        item.className = 'dropdown-item selected';
        item.textContent = site;
        item.dataset.value = site;
        
        item.addEventListener('click', () => {
            item.classList.toggle('selected');
            updateDashboard();
        });
        locContainer.appendChild(item);
    });

    categories.forEach(cat => {
        const item = document.createElement('div');
        item.className = 'dropdown-item selected';
        item.textContent = cat;
        item.dataset.value = cat;
        
        item.addEventListener('click', () => {
            item.classList.toggle('selected');
            syncSelectAllToggles();
            updateDashboard();
        });
        catContainer.appendChild(item);
    });

    syncSelectAllToggles();
}

function syncSelectAllToggles() {
    const catToggle = document.getElementById('category-select-all-toggle');
    const locToggle = document.getElementById('location-select-all-toggle');

    if (catToggle) {
        const items = Array.from(document.querySelectorAll('#category-filters-list .dropdown-item'));
        catToggle.checked = items.length > 0 && items.every(el => el.classList.contains('selected'));
    }

    if (locToggle) {
        const items = Array.from(document.querySelectorAll('#location-filters-list .dropdown-item'));
        locToggle.checked = items.length > 0 && items.every(el => el.classList.contains('selected'));
    }
}


function updateDashboard() {
    if (window.appDatasets.length === 0) {
        noDataPlaceholder.style.display = 'flex';
        const viewControls = document.getElementById('view-controls');
        if (viewControls) viewControls.style.display = 'none';
        document.getElementById('charts-wrapper').style.display = 'none';
        
        document.getElementById('location-filters-list').innerHTML = '';
        document.getElementById('category-filters-list').innerHTML = '';
        
        document.getElementById('main-diversion-rate').textContent = "--%";
        document.getElementById('main-total-waste').textContent = "-- t";
        document.getElementById('main-landfill').textContent = "-- t";
        document.getElementById('sidebar-insights-container').innerHTML = `<div class="insight-item"><span class="insight-number">1</span><p>Awaiting data upload to generate insights.</p></div>`;
        return;
    }

    let allWasteRows = [];
    let wasteHeaders = [];
    let allSites = new Set();
    let allCats = new Set();
    
    let allSalesRows = [];
    let salesHeaders = [];

    window.appDatasets.forEach(d => {
        if (d.type === 'waste') {
            allWasteRows = allWasteRows.concat(d.rawRows);
            if (wasteHeaders.length === 0) wasteHeaders = d.rawHeaders;
            d.uniqueSites.forEach(s => allSites.add(s));
            d.uniqueCategories.forEach(c => allCats.add(c));
        } else if (d.type === 'sales') {
            allSalesRows = allSalesRows.concat(d.rawRows);
            if (salesHeaders.length === 0) salesHeaders = d.rawHeaders;
        }
    });

    // Filters are only built on first load so we do not wipe out the user's current selections
    // every time a dataset is added or removed.
    const locContainer = document.getElementById('location-filters-list');
    const catContainer = document.getElementById('category-filters-list');
    if (locContainer.children.length === 0 && allSites.size > 0) {
        populateFilters(Array.from(allSites), Array.from(allCats));
    }

    const selectedSites = Array.from(document.querySelectorAll('#location-filters-list .dropdown-item.selected')).map(el => el.dataset.value);
    const selectedCats = Array.from(document.querySelectorAll('#category-filters-list .dropdown-item.selected')).map(el => el.dataset.value);

    let filteredWasteData = { timeline: [], metrics: { diversionRate: 0, totalWaste: 0, totalLandfill: 0 }, uniqueCategories: [], topLocations: [], insights: [] };
    if (allWasteRows.length > 0) {
        filteredWasteData = processData(allWasteRows, wasteHeaders, { sites: selectedSites, categories: selectedCats });
    }

    let processedSalesData = { categories: [], totalRevenue: 0 };
    if (allSalesRows.length > 0) {
        processedSalesData = processSalesData(allSalesRows, salesHeaders);
    }
    
    noDataPlaceholder.style.display = 'none';
    document.getElementById('charts-wrapper').style.display = 'grid';
    renderCharts(filteredWasteData.timeline, filteredWasteData.uniqueCategories, filteredWasteData.topLocations, processedSalesData);

    // Store the latest aggregate values globally because both the sidebar and maximized overlay
    // read from the same source of truth after chart rerenders.
    const metricsData = filteredWasteData.metrics || { diversionRate: 0, totalWaste: 0, totalLandfill: 0 };
    window.currentMetricsData = metricsData;
    
    document.getElementById('main-diversion-rate').textContent = metricsData.diversionRate.toFixed(1) + "%";
    document.getElementById('main-total-waste').textContent = Math.round(metricsData.totalWaste).toLocaleString() + " t";
    document.getElementById('main-landfill').textContent = Math.round(metricsData.totalLandfill).toLocaleString() + " t";

    
    window.currentInsightsMap = {
        'chart-category': [],
        'chart-disposition-detailed': [],
        'chart-disposition-diverted': [],
        'chart-locations': [],
        'chart-sales': []
    };
    
    if (processedSalesData.insights) {
        Object.assign(window.currentInsightsMap, processedSalesData.insights);
    }
    if (filteredWasteData.insights) {
        for (const [k, v] of Object.entries(filteredWasteData.insights)) {
            if (!window.currentInsightsMap[k]) window.currentInsightsMap[k] = [];
            window.currentInsightsMap[k] = window.currentInsightsMap[k].concat(v);
        }
    }
    
    renderInsights();
    renderSidebarInsights();
}

function renderSidebarInsights() {
    const container = document.getElementById('sidebar-insights-container');
    if (!container) return;

    const map = window.currentInsightsMap || {};
    const chartPriority = [
        'chart-disposition-diverted',
        'chart-disposition-detailed',
        'chart-locations',
        'chart-category',
        'chart-sales'
    ];

    const picked = [];
    const seen = new Set();
    const seenTexts = [];

    const normalizeInsight = (t) => {
        if (typeof t !== 'string') return '';
        return t
            .trim()
            .replace(/\s+/g, ' ')
            .replace(/[“”]/g, '"')
            .replace(/[‘’]/g, "'")
            .toLowerCase();
    };

    const isNearDuplicate = (normalized) => {
        if (!normalized) return false;
        for (const prev of seenTexts) {
            if (prev === normalized) return true;
            if (prev.includes(normalized) || normalized.includes(prev)) return true;
        }
        return false;
    };

    for (const k of chartPriority) {
        const arr = map[k] || [];
        for (const t of arr) {
            if (typeof t === 'string' && t.trim().length > 0) {
                const cleaned = t.trim();
                const key = normalizeInsight(cleaned);
                if (key && !seen.has(key) && !isNearDuplicate(key)) {
                    seen.add(key);
                    seenTexts.push(key);
                    picked.push(cleaned);
                    break;
                }
            }
        }
    }

    const displayInsights = picked.slice(0, 4);
    if (displayInsights.length === 0) {
        container.innerHTML = `<div class="insight-item"><span class="insight-number">1</span><p>Awaiting data extraction parameters or threshold events to generate insights.</p></div>`;
        return;
    }

    container.innerHTML = displayInsights
        .map((text, i) => `<div class="insight-item"><span class="insight-number">${i + 1}</span><p>${text}</p></div>`)
        .join('');
}

let currentInsightsMap = {
    'chart-category': [],
    'chart-disposition-detailed': [],
    'chart-disposition-diverted': [],
    'chart-locations': [],
    'chart-sales': []
};
window.currentInsightsMap = currentInsightsMap;

let currentlyMaximized = null;

function renderInsights() {
    document.querySelectorAll('.maximized-insights-overlay').forEach(el => el.remove());
    if (!currentlyMaximized) return;

    const maxContainer = document.getElementById(currentlyMaximized);
    const overlay = document.createElement('div');
    overlay.className = 'maximized-insights-overlay';
    
    const m = window.currentMetricsData || { diversionRate: 0, totalWaste: 0, totalLandfill: 0 };
    
    let htmlBlock = `
      <div class="panel-section stats-section">
        <div class="panel-header">Key Metrics</div>
        
        <div class="metric-card">
          <div class="metric-value">
            <span class="metric-big">${m.diversionRate.toFixed(1)}%</span>
            <span class="metric-sub">Avg Diversion Rate</span>
          </div>
        </div>

        <div class="metric-card">
          <div class="metric-value">
            <span class="metric-big">${Math.round(m.totalWaste).toLocaleString()} t</span>
            <span class="metric-sub">Total Waste Generated (tonnes)</span>
          </div>
        </div>

        <div class="metric-card" style="border-bottom: none; margin-bottom: 24px;">
          <div class="metric-value">
            <span class="metric-big">${Math.round(m.totalLandfill).toLocaleString()} t</span>
            <span class="metric-sub">Landfill Waste (tonnes)</span>
          </div>
        </div>
      </div>

      <div class="panel-section insights-section">
        <div class="panel-header" style="margin-top: 10px;">AUTOMATED INSIGHTS</div>
        <div id="insights-container" style="padding-top: 10px;">
    `;

    let displayInsights = window.currentInsightsMap[currentlyMaximized] || [];
    displayInsights = displayInsights.slice(0, 2); 
    
    if (displayInsights.length > 0) {
        displayInsights.forEach((insightText, i) => {
            htmlBlock += `
            <div class="insight-item">
              <span class="insight-number">${i + 1}</span>
              <p>${insightText}</p>
            </div>
            `;
        });
    } else {
        htmlBlock += `
            <div class="insight-item">
              <span class="insight-number">1</span>
              <p style="color: var(--bbc-text-secondary); font-size: 13px;">Awaiting data extraction parameters or threshold events to generate insights.</p>
            </div>
        `;
    }
    
    htmlBlock += `</div></div>`;
    
    overlay.innerHTML = htmlBlock;
    maxContainer.appendChild(overlay);
}

// Maximized charts reuse the same Chart.js instances, so we resize after the CSS transition
// finishes to prevent clipped canvases.
document.querySelectorAll('.maximize-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
        const container = e.target.closest('.canvas-container');
        const targetId = container.id;
        
        if (container.classList.contains('maximized-chart')) {
            container.classList.remove('maximized-chart');
            document.getElementById('modal-backdrop').classList.remove('active');
            currentlyMaximized = null;
        } else {
            document.querySelectorAll('.canvas-container').forEach(el => el.classList.remove('maximized-chart'));
            container.classList.add('maximized-chart');
            document.getElementById('modal-backdrop').classList.add('active');
            currentlyMaximized = targetId;
        }
        
        renderInsights();
        
        const chartMap = {
            'chart-category': window.char1,
            'chart-disposition-diverted': window.char2,
            'chart-locations': window.char3,
            'chart-disposition-detailed': window.char4,
            'chart-sales': window.char5
        };
        const activeChart = chartMap[currentlyMaximized || targetId];
        if (activeChart) {
            setTimeout(() => activeChart.resize(), 10);
            setTimeout(() => activeChart.resize(), 300);
        }
    });
});

document.getElementById('modal-backdrop').addEventListener('click', () => {
    document.querySelectorAll('.canvas-container').forEach(el => el.classList.remove('maximized-chart'));
    document.getElementById('modal-backdrop').classList.remove('active');
    
    const previousTarget = currentlyMaximized;
    currentlyMaximized = null;
    renderInsights();
    
    const chartMap = {
        'chart-category': window.char1,
        'chart-disposition-diverted': window.char2,
        'chart-locations': window.char3,
        'chart-disposition-detailed': window.char4,
        'chart-sales': window.char5
    };
    const activeChart = chartMap[previousTarget];
    if (activeChart) {
        setTimeout(() => activeChart.resize(), 10);
        setTimeout(() => activeChart.resize(), 300);
    }
});

document.getElementById('category-dropdown-btn').addEventListener('click', () => {
    const list = document.getElementById('category-filters-list');
    list.style.display = list.style.display === 'none' ? 'block' : 'none';
});

document.getElementById('location-dropdown-btn').addEventListener('click', () => {
    const list = document.getElementById('location-filters-list');
    list.style.display = list.style.display === 'none' ? 'block' : 'none';
});

document.getElementById('filters-collapse-btn').addEventListener('click', () => {
    const section = document.querySelector('.filter-section');
    const body = document.getElementById('filters-body');
    const btn = document.getElementById('filters-collapse-btn');
    if (!section || !body || !btn) return;

    const willExpand = body.style.display === 'none';
    body.style.display = willExpand ? 'block' : 'none';
    section.classList.toggle('filters-collapsed', !willExpand);

    btn.setAttribute('aria-expanded', String(willExpand));
    const label = btn.querySelector('.filters-collapse-label');
    if (label) label.textContent = willExpand ? 'Collapse' : 'Expand';

    if (!willExpand) {
        const catList = document.getElementById('category-filters-list');
        const locList = document.getElementById('location-filters-list');
        if (catList) catList.style.display = 'none';
        if (locList) locList.style.display = 'none';
    }
});

document.getElementById('category-select-all-toggle').addEventListener('click', (e) => {
    e.stopPropagation();
});

document.getElementById('location-select-all-toggle').addEventListener('click', (e) => {
    e.stopPropagation();
});

document.getElementById('category-select-all-toggle').addEventListener('change', (e) => {
    const isOn = e.target.checked;
    document.querySelectorAll('#category-filters-list .dropdown-item')
        .forEach(el => el.classList.toggle('selected', isOn));
    syncSelectAllToggles();
    updateDashboard();
});

document.getElementById('location-select-all-toggle').addEventListener('change', (e) => {
    const isOn = e.target.checked;
    document.querySelectorAll('#location-filters-list .dropdown-item')
        .forEach(el => el.classList.toggle('selected', isOn));
    syncSelectAllToggles();
    updateDashboard();
});
