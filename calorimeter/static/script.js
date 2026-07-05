const PARAM_META = {}; // populated from checkbox labels at runtime
let dashboardData = null; // cached payload from /upload or /data

const folderInput = document.getElementById('folderInput');
const folderLabel = document.getElementById('folderLabel');
const uploadStatus = document.getElementById('uploadStatus');
const emptyState = document.getElementById('emptyState');
const kpiRow = document.getElementById('kpiRow');
const tablesContainer = document.getElementById('tablesContainer');
const exportXlsxBtn = document.getElementById('exportXlsx');
const exportPdfBtn = document.getElementById('exportPdf');
const themeToggle = document.getElementById('themeToggle');

// Build param meta map from the checkbox list rendered by the template
document.querySelectorAll('.param-item').forEach(item => {
  const input = item.querySelector('.param-check');
  const label = item.querySelector('span').textContent;
  PARAM_META[input.value] = { label };
});

// ---------- Theme ----------
themeToggle.addEventListener('click', () => {
  const html = document.documentElement;
  const next = html.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
  html.setAttribute('data-theme', next);
  themeToggle.innerHTML = next === 'dark' ? '&#9788;' : '&#9789;';
});

// ---------- Folder upload ----------
folderInput.addEventListener('change', async (e) => {
  const files = Array.from(e.target.files).filter(f => f.name.toLowerCase().endsWith('.pdf'));
  if (files.length === 0) {
    uploadStatus.textContent = 'No PDF files found in the selected folder.';
    uploadStatus.classList.add('error');
    return;
  }

  folderLabel.textContent = `${files.length} PDF report(s) selected`;
  uploadStatus.classList.remove('error');
  uploadStatus.textContent = 'Parsing reports…';

  const formData = new FormData();
  files.forEach(f => {
    const relPath = f.webkitRelativePath || f.name;
    formData.append('files', f, relPath);
  });

  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const json = await res.json();
    if (json.error) throw new Error(json.error);

    dashboardData = json.data;
    uploadStatus.textContent = `Parsed ${json.count} report(s) successfully.` +
      (json.errors.length ? ` (${json.errors.length} failed)` : '');
    if (json.errors.length) uploadStatus.classList.add('error');

    exportXlsxBtn.disabled = false;
    exportPdfBtn.disabled = false;
    render();
  } catch (err) {
    uploadStatus.textContent = 'Error parsing reports: ' + err.message;
    uploadStatus.classList.add('error');
  }
});

// ---------- Parameter checkboxes ----------
document.getElementById('selAll').addEventListener('click', () => {
  document.querySelectorAll('.param-check').forEach(c => c.checked = true);
  render();
});
document.getElementById('selNone').addEventListener('click', () => {
  document.querySelectorAll('.param-check').forEach(c => c.checked = false);
  render();
});
document.querySelectorAll('.param-check').forEach(c => c.addEventListener('change', render));

function getSelectedParams() {
  return Array.from(document.querySelectorAll('.param-check'))
    .filter(c => c.checked)
    .map(c => c.value);
}

// ---------- Rendering ----------
function fmt(val, key) {
  if (val === null || val === undefined) return null;
  const decimals = {
    rpm_actual: 0, cooling_capacity: 1, input_power: 2, cop: 3, current: 3,
    inv_eff: 2, inv_loss: 2, power_factor: 3, shell_top_temp: 1,
    discharge_temp: 1, mass_flow: 0, vol_eff: 3, isen_eff: 3
  };
  const d = decimals[key] !== undefined ? decimals[key] : 2;
  return Number(val).toFixed(d);
}

const STAT_LABELS = { cop: 'COP (W/W)', cooling_capacity: 'Cooling Capacity (W)', input_power: 'Input Power (W)' };
const STAT_DECIMALS = { cop: 3, cooling_capacity: 1, input_power: 2 };

function fmtStat(val, key) {
  if (val === null || val === undefined) return '—';
  return Number(val).toFixed(STAT_DECIMALS[key] ?? 2);
}

function renderTierStatsHeader(tierBlock) {
  const row = document.createElement('div');
  row.className = 'tier-stats-row';
  ['cop', 'cooling_capacity', 'input_power'].forEach(key => {
    const s = tierBlock.stats[key];
    const card = document.createElement('div');
    card.className = 'tier-stat-card';
    if (!s) {
      card.innerHTML = `<span class="stat-label">${STAT_LABELS[key]}</span><div class="stat-main-row"><span class="stat-range">—</span></div>`;
    } else {
      card.innerHTML = `
        <span class="stat-label">${STAT_LABELS[key]}</span>
        <div class="stat-main-row">
          <span class="stat-range">${fmtStat(s.min, key)} – ${fmtStat(s.max, key)}</span>
          <span class="stat-std">σ = <b>${fmtStat(s.std, key)}</b></span>
        </div>`;
    }
    row.appendChild(card);
  });
  return row;
}

function render() {
  if (!dashboardData) return;
  const selected = getSelectedParams();

  emptyState.style.display = 'none';
  kpiRow.style.display = 'flex';
  document.getElementById('chartsSection').style.display = 'block';

  const s = dashboardData.summary;
  document.getElementById('kpiTotal').textContent = s.total_reports;
  document.getElementById('kpiIds').textContent = s.unique_ids;
  document.getElementById('kpiTiers').textContent = s.rpm_tiers;
  document.getElementById('kpiAvgCop').textContent = s.avg_cop ?? '—';
  document.getElementById('kpiMaxCop').textContent = s.max_cop ?? '—';

  tablesContainer.innerHTML = '';

  dashboardData.tiers.forEach(tierBlock => {
    const block = document.createElement('div');
    block.className = 'tier-block';

    const title = document.createElement('h3');
    title.className = 'tier-title';
    title.textContent = `${tierBlock.tier} RPM Test Point`;
    block.appendChild(title);

    // per-tier headline stats header (range + std dev for COP / Cooling / Input Power)
    block.appendChild(renderTierStatsHeader(tierBlock));

    const wrap = document.createElement('div');
    wrap.className = 'table-wrap';

    const table = document.createElement('table');
    table.className = 'compare-table';

    // determine best COP id for highlight
    let bestId = null;
    if (selected.includes('cop')) {
      let bestVal = -Infinity;
      tierBlock.ids.forEach(id => {
        const v = tierBlock.rows[id].cop;
        if (v !== null && v !== undefined && v > bestVal) { bestVal = v; bestId = id; }
      });
    }

    // ---- Transposed: Parameter rows x Compressor ID columns ----
    const thead = document.createElement('thead');
    const headRow = document.createElement('tr');
    headRow.innerHTML = '<th class="param-col">Parameter</th>' +
      tierBlock.ids.map(id => `<th>${id}</th>`).join('');
    thead.appendChild(headRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    selected.forEach(k => {
      const tr = document.createElement('tr');
      let cells = `<td class="param-col">${PARAM_META[k] ? PARAM_META[k].label : k}</td>`;
      tierBlock.ids.forEach(id => {
        const raw = tierBlock.rows[id][k];
        const display = fmt(raw, k);
        const isBest = (k === 'cop' && id === bestId);
        if (display === null) {
          cells += `<td class="na">—</td>`;
        } else {
          cells += `<td class="${isBest ? 'best' : ''}">${display}</td>`;
        }
      });
      tr.innerHTML = cells;
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);

    wrap.appendChild(table);
    block.appendChild(wrap);
    tablesContainer.appendChild(block);
  });

  renderCharts();
}

// ---------- Sample-wise trend charts (separate section) ----------
let chartInstances = [];

function destroyCharts() {
  chartInstances.forEach(c => c.destroy());
  chartInstances = [];
}

function renderCharts() {
  if (!dashboardData) return;
  destroyCharts();
  const container = document.getElementById('chartsContainer');
  container.innerHTML = '';

  const chartParams = [
    { key: 'cop', label: 'COP (W/W)', color: '#3d8bfd' },
    { key: 'cooling_capacity', label: 'Cooling Capacity (W)', color: '#4caf7d' },
    { key: 'input_power', label: 'Input Power (W)', color: '#e0a83e' },
  ];

  dashboardData.tiers.forEach(tierBlock => {
    const tierBlockEl = document.createElement('div');
    tierBlockEl.className = 'chart-tier-block';

    const heading = document.createElement('h4');
    heading.className = 'chart-tier-heading';
    heading.textContent = `${tierBlock.tier} RPM Test Point`;
    tierBlockEl.appendChild(heading);

    const grid = document.createElement('div');
    grid.className = 'chart-grid';

    chartParams.forEach(cp => {
      const card = document.createElement('div');
      card.className = 'chart-card';
      const canvas = document.createElement('canvas');
      card.appendChild(canvas);
      grid.appendChild(card);

      const values = tierBlock.ids.map(id => tierBlock.rows[id][cp.key]);

      const chart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
          labels: tierBlock.ids,
          datasets: [{
            label: cp.label,
            data: values,
            backgroundColor: cp.color + 'cc',
            borderRadius: 4,
            maxBarThickness: 34
          }]
        },
        options: {
          responsive: true,
          plugins: {
            legend: { display: false },
            title: { display: true, text: cp.label, font: { size: 12 } }
          },
          scales: {
            x: { ticks: { font: { size: 9 }, maxRotation: 45, minRotation: 45 } },
            y: { beginAtZero: false, ticks: { font: { size: 9 } } }
          }
        }
      });
      chartInstances.push(chart);
    });

    tierBlockEl.appendChild(grid);
    container.appendChild(tierBlockEl);
  });
}

// ---------- Export ----------
async function doExport(kind) {
  const selected = getSelectedParams();
  const url = kind === 'xlsx' ? '/export/xlsx' : '/export/pdf';
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ params: selected })
  });
  if (!res.ok) { alert('Export failed.'); return; }
  const blob = await res.blob();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = kind === 'xlsx' ? 'calorimeter_comparison.xlsx' : 'calorimeter_comparison.pdf';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

exportXlsxBtn.addEventListener('click', () => doExport('xlsx'));
exportPdfBtn.addEventListener('click', () => doExport('pdf'));
