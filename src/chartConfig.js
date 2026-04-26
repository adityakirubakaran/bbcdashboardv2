import { Chart, registerables } from 'chart.js';
Chart.register(...registerables);

// Resize animations make filter changes feel like the canvas is zooming. We disable just the
// resize transition so data updates stay responsive without removing chart draw animations.
Chart.defaults.transitions = Chart.defaults.transitions || {};
if (!Chart.defaults.transitions.resize) Chart.defaults.transitions.resize = { animation: {} };
Chart.defaults.transitions.resize.animation.duration = 0;

window.char1 = window.char1 || null;
window.char2 = window.char2 || null;
window.char3 = window.char3 || null;
window.char4 = window.char4 || null;
window.char5 = window.char5 || null;

const CATEGORY_COLORS = [
  '#3A3532', '#615A55', '#9D6C5B', '#8F978E', '#C0AFA2', '#8D837C', '#D3C5B5', '#6B746A', '#A49182', '#1D1C1B'
];

const tonnesLabel = (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString(undefined, { maximumFractionDigits: 1 })} t`;
const tonnesLabelX = (ctx) => `${ctx.dataset.label}: ${ctx.parsed.x.toLocaleString(undefined, { maximumFractionDigits: 1 })} t`;

export function renderCharts(timelineData, uniqueCategories, topLocations, salesData = null) {
  const labels = timelineData.map(d => d.month);

  const ctx1 = document.getElementById('category-chart').getContext('2d');

  // Category keys can disappear after filtering, so we derive datasets from the filtered timeline
  // instead of trusting the original header list.
  let activeCategories = new Set();
  timelineData.forEach(m => {
      Object.keys(m.categoryVolumes).forEach(k => activeCategories.add(k));
  });
  const catArray = Array.from(activeCategories);

  const categoryDatasets = catArray.map((cat, idx) => {
      return {
          label: cat,
          data: timelineData.map(d => d.categoryVolumes[cat] || 0),
          backgroundColor: CATEGORY_COLORS[idx % CATEGORY_COLORS.length],
          stack: 'combined'
      };
  });

  if (window.char1) window.char1.destroy();
  window.char1 = new Chart(ctx1, {
      type: 'bar',
      data: { labels, datasets: categoryDatasets },
      options: {
          maintainAspectRatio: false,
          responsive: true,
          plugins: {
              legend: { position: 'right', labels: { font: { family: 'Inter', size: 11 }, usePointStyle: true, boxWidth: 8 } },
              tooltip: { callbacks: { label: tonnesLabel } }
          },
          scales: { x: { stacked: true, grid: { display: false } }, y: { stacked: true, title: { display: true, text: 'Tonnes' } } }
      }
  });

  const ctx2 = document.getElementById('disposition-chart').getContext('2d');

  const divertedData = timelineData.map(d => d.diverted);
  const landfillData = timelineData.map(d => d.landfill);

  if (window.char2) window.char2.destroy();
  window.char2 = new Chart(ctx2, {
      type: 'line',
      data: {
          labels,
          datasets: [
              {
                  label: 'Diverted from Landfill',
                  data: divertedData,
                  borderColor: '#8F978E',
                  backgroundColor: 'rgba(143, 151, 142, 0.4)',
                  fill: true,
                  tension: 0.4
              },
              {
                  label: 'Landfill',
                  data: landfillData,
                  borderColor: '#9D6C5B',
                  backgroundColor: 'rgba(157, 108, 91, 0.4)',
                  fill: true,
                  tension: 0.4
              }
          ]
      },
      options: {
          maintainAspectRatio: false,
          animation: false,
          responsive: true,
          plugins: {
              legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 12 }, usePointStyle: true, boxWidth: 10 } },
              tooltip: { callbacks: { label: tonnesLabel } }
          },
          scales: {
              x: { grid: { display: false } },
              y: { title: { display: true, text: 'Tonnes' }, min: 0 }
          }
      }
  });

  const ctx4 = document.getElementById('disposition-detailed-chart').getContext('2d');

  const lineAd = timelineData.map(d => d.rawAd);
  const lineInc = timelineData.map(d => d.rawIncineration);
  const lineRecycled = timelineData.map(d => d.rawRecycled);
  const lineLandfill = timelineData.map(d => d.rawLandfill);

  if (window.char4) window.char4.destroy();
  window.char4 = new Chart(ctx4, {
      type: 'bar',
      data: {
          labels,
          datasets: [
              { label: 'AD', data: lineAd, backgroundColor: '#6B746A', stack: 'dStack' },
              { label: 'Incineration', data: lineInc, backgroundColor: '#9D6C5B', stack: 'dStack' },
              { label: 'Recycled', data: lineRecycled, backgroundColor: '#C0AFA2', stack: 'dStack' },
              { label: 'Landfill', data: lineLandfill, backgroundColor: '#3A3532', stack: 'dStack' }
          ]
      },
      options: {
          maintainAspectRatio: false,
          responsive: true,
          plugins: {
              legend: { position: 'bottom', labels: { font: { family: 'Inter', size: 11 }, usePointStyle: true } },
              tooltip: { callbacks: { label: tonnesLabel } }
          },
          scales: {
              x: { stacked: true, grid: { display: false } },
              y: { stacked: true, title: { display: true, text: 'Tonnes' } }
          }
      }
  });

  const ctx3 = document.getElementById('location-chart').getContext('2d');

  const siteLabels = topLocations.map(l => l.site);
  const siteDiverted = topLocations.map(l => l.diverted);
  const siteLandfill = topLocations.map(l => l.landfill);

  if (window.char3) window.char3.destroy();
  window.char3 = new Chart(ctx3, {
      type: 'bar',
      data: {
          labels: siteLabels,
          datasets: [
              { label: 'Diverted', data: siteDiverted, backgroundColor: '#8F978E', stack: 'stack1' },
              { label: 'Landfill', data: siteLandfill, backgroundColor: '#9D6C5B', stack: 'stack1' }
          ]
      },
      options: {
          maintainAspectRatio: false,
          indexAxis: 'y',
          responsive: true,
          plugins: {
              legend: { display: false },
              tooltip: { callbacks: { label: tonnesLabelX } }
          },
          scales: {
              x: { stacked: true, title: { display: true, text: 'Total Tonnes' } },
              y: { stacked: true, grid: { display: false }, ticks: { font: { size: 10 } } }
          }
      }
  });
  const ctx5 = document.getElementById('sales-chart') ? document.getElementById('sales-chart').getContext('2d') : null;
  if (ctx5 && salesData && salesData.categories.length > 0) {
      if (window.char5) window.char5.destroy();

      const salesLabels = salesData.categories.map(c => c.category);
      const salesRevs = salesData.categories.map(c => c.revenue);

      window.char5 = new Chart(ctx5, {
          type: 'bar',
          data: {
              labels: salesLabels,
              datasets: [{
                  label: 'Revenue (£)',
                  data: salesRevs,
                  backgroundColor: CATEGORY_COLORS.slice(0, salesLabels.length)
              }]
          },
          options: {
              maintainAspectRatio: false,
              responsive: true,
              plugins: {
                  legend: { display: false }
              },
              scales: {
                  x: { grid: { display: false } },
                  y: { title: { display: true, text: 'Revenue (£)' }, beginAtZero: true }
              }
          }
      });
  } else if (ctx5 && window.char5) {
      window.char5.destroy();
      window.char5 = null;
  }
}
