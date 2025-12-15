/* ==================== Financial Dashboard - Standalone App ==================== */

// Global Variables
let dashboardData = null;
let charts = {};
let currentMonth = null;
let settings = {
  apiUrl: '',
  currency: 'SAR',
  darkMode: false
};

/* ==================== Initialization ==================== */
document.addEventListener('DOMContentLoaded', function() {
  loadSettings();
  initializeNavigation();
  loadDashboardData();
});

function initializeNavigation() {
  // Navigation click handlers
  document.querySelectorAll('.nav-item[data-section]').forEach(item => {
    item.addEventListener('click', function(e) {
      e.preventDefault();
      const section = this.getAttribute('data-section');
      showSection(section);
    });
  });

  // Chart range buttons
  document.querySelectorAll('.chart-btn').forEach(btn => {
    btn.addEventListener('click', function() {
      const range = this.getAttribute('data-range');
      document.querySelectorAll('.chart-btn').forEach(b => b.classList.remove('active'));
      this.classList.add('active');
      updateChartRange(range);
    });
  });
}

/* ==================== Settings ==================== */
function loadSettings() {
  const saved = localStorage.getItem('dashboardSettings');
  if (saved) {
    settings = JSON.parse(saved);
    const apiUrlInput = document.getElementById('apiUrl');
    if (apiUrlInput) apiUrlInput.value = settings.apiUrl || '';
    document.getElementById('currency').value = settings.currency || 'SAR';
    document.getElementById('darkMode').checked = settings.darkMode || false;
    if (settings.darkMode) {
      document.documentElement.setAttribute('data-theme', 'dark');
    }
  }
}

function saveSettings() {
  const apiUrlInput = document.getElementById('apiUrl');
  if (apiUrlInput) settings.apiUrl = apiUrlInput.value.trim();
  settings.currency = document.getElementById('currency').value;
  settings.darkMode = document.getElementById('darkMode').checked;
  localStorage.setItem('dashboardSettings', JSON.stringify(settings));
  closeSettings();
  showToast('success', 'Settings Saved', 'Your preferences have been saved');
  loadDashboardData(); // Reload with new settings
}

function showSettings() {
  document.getElementById('settingsModal').classList.add('active');
}

function closeSettings() {
  document.getElementById('settingsModal').classList.remove('active');
}

function toggleDarkMode() {
  const isDark = document.getElementById('darkMode').checked;
  document.documentElement.setAttribute('data-theme', isDark ? 'dark' : 'light');
}

/* ==================== Data Loading ==================== */
function loadDashboardData() {
  showLoading(true);

  // If API URL is configured, try to fetch live data using JSONP
  if (settings.apiUrl) {
    fetchWithJSONP(settings.apiUrl)
      .then(data => {
        if (data && !data.error) {
          dashboardData = data;
          currentMonth = data.latestMonth;
          updateAllDisplays();
          showLoading(false);
          showToast('success', 'Live Data', 'Dashboard updated with live data from Google Sheets');
        } else {
          console.warn('API returned error:', data?.error);
          loadDemoData();
          showToast('warning', 'Using Demo Data', data?.error || 'Could not fetch live data');
        }
      })
      .catch(error => {
        console.error('API Error:', error);
        loadDemoData();
        showToast('warning', 'Using Demo Data', 'API unavailable - showing demo data');
      });
  } else {
    // No API URL configured, use demo data
    loadDemoData();
  }
}

/**
 * Fetch data using JSONP (bypasses CORS restrictions)
 */
function fetchWithJSONP(apiUrl, timeout = 15000) {
  return new Promise((resolve, reject) => {
    // Generate unique callback name
    const callbackName = 'jsonpCallback_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);

    // Build URL with JSONP parameters
    let url = apiUrl;
    if (url.endsWith('/')) {
      url = url.slice(0, -1);
    }
    const separator = url.includes('?') ? '&' : '?';
    url += separator + 'action=getData&callback=' + callbackName;

    // Create script element
    const script = document.createElement('script');
    script.src = url;

    // Timeout handler
    const timeoutId = setTimeout(() => {
      cleanup();
      reject(new Error('Request timed out'));
    }, timeout);

    // Cleanup function
    function cleanup() {
      clearTimeout(timeoutId);
      delete window[callbackName];
      if (script.parentNode) {
        script.parentNode.removeChild(script);
      }
    }

    // Define callback function
    window[callbackName] = function(data) {
      cleanup();
      resolve(data);
    };

    // Error handler
    script.onerror = function() {
      cleanup();
      reject(new Error('Failed to load script'));
    };

    // Add script to page (triggers the request)
    document.head.appendChild(script);
  });
}

function loadDemoData() {
  // Demo data based on actual spreadsheet
  dashboardData = {
    latestMonth: 'NOV-2025',
    months: ['JAN-2025', 'FEB-2025', 'MAR-2025', 'APR-2025', 'MAY-2025', 'JUN-2025', 'JUL-2025', 'AUG-2025', 'SEP-2025', 'OCT-2025', 'NOV-2025'],
    kpi: {
      bankBalance: 2498320.62,
      bankChange: 1.97,
      outstanding: 2179683.50,
      outstandingChange: 5.66,
      advances: 540959.39,
      advancesChange: 7.22,
      suspense: -17267.64,
      suspenseChange: 0
    },
    ytd: {
      received: 28577792.19,
      payments: 27400972.16,
      netFlow: 1176820.03,
      months: 11
    },
    banks: {
      labels: ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV'],
      balance: [1143292.67, 1476523.89, 1823456.12, 2156789.34, 2489012.56, 2821345.78, 3154678.90, 3487901.12, 3820234.34, 3850313.61, 2498320.62],
      received: [2597981, 2800000, 2600000, 2700000, 2900000, 2400000, 2600000, 2800000, 2700000, 2500000, 2477792],
      payments: [2264658, 2500000, 2400000, 2200000, 2500000, 2200000, 2400000, 2600000, 2500000, 2650000, 2400972]
    },
    outstanding: {
      labels: ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV'],
      total: [1795869.49, 1589925.60, 1507837.21, 1716100.16, 2287954.21, 2189458.69, 2399787.77, 2026772.41, 2023784.28, 2063023.08, 2179683.50],
      salesmen: [
        { name: 'company sales', value: 976412.87, trend: -7.58 },
        { name: 'arshad', value: 486637.40, trend: 32.79 },
        { name: 'dhiya', value: 270754.61, trend: 10.23 },
        { name: 'Shameer riyadh', value: 92907.92, trend: 27.87 },
        { name: 'khalid', value: 89837.02, trend: -9.26 },
        { name: 'nidheesh', value: 79345.24, trend: 3.27 },
        { name: 'akmal', value: 34121.05, trend: -25.50 },
        { name: 'samir dmm', value: 30912.40, trend: 100 },
        { name: 'bassam', value: 22733.87, trend: -9.85 },
        { name: 'al ahsa-4', value: 20589.67, trend: 100 },
        { name: 'jaseel', value: 19951.36, trend: -2.70 },
        { name: 'shafi', value: 18398.82, trend: 6.60 },
        { name: 'ashik', value: 17807.28, trend: -36.24 },
        { name: 'samir', value: 15016.49, trend: 65.23 },
        { name: 'shareef', value: 4257.50, trend: 0 }
      ]
    },
    advances: {
      labels: ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV'],
      opening: [163800.84, 479608.52, 489038.28, 491564.28, 541048.28, 472788.43, 504240.98, 452575.39, 473840.39, 504511.39, 540959.39],
      given: [629615.36, 813275.72, 653526.88, 752494.88, 615975.18, 678880.28, 577549.10, 618079.10, 681421.10, 0, 0],
      settled: [313807.68, 803845.96, 651000.88, 703010.88, 684235.03, 647427.73, 629214.69, 596814.10, 650750.10, 0, 0],
      closing: [479608.52, 489038.28, 491564.28, 541048.28, 472788.43, 504240.98, 452575.39, 473840.39, 504511.39, 540959.39, 540959.39]
    },
    suspense: {
      labels: ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV'],
      balance: [-17267.64, -17267.64, -17267.64, -17267.64, -17267.64, -17267.64, -17267.64, -17267.64, -17267.64, -17267.64, -17267.64]
    },
    bankAccounts: [
      { name: 'Alrajhi-1097', balance: 210693.67, change: 0 },
      { name: 'Alrajhi-new', balance: 564500.97, change: 0 },
      { name: 'SNB Al Ahsa Branch', balance: 1562570.61, change: 0 },
      { name: 'Albilad', balance: 160054.48, change: 0 },
      { name: 'Albilad USD', balance: 133.57, change: 0 }
    ]
  };

  currentMonth = dashboardData.latestMonth;
  populateMonthSelector(dashboardData.months);
  updateDashboard(dashboardData);
  showLoading(false);
}

/* ==================== Dashboard Updates ==================== */
function updateDashboard(data) {
  updateKPIs(data.kpi);
  updateYTD(data.ytd);
  updateCharts(data);
  updateTables(data);
}

function updateKPIs(kpi) {
  // Bank Balance
  document.getElementById('kpi-bank-balance').textContent = formatCurrency(kpi.bankBalance);
  updateChangeIndicator('kpi-bank-change', kpi.bankChange, true);

  // Outstanding (lower is better)
  document.getElementById('kpi-outstanding').textContent = formatCurrency(kpi.outstanding);
  updateChangeIndicator('kpi-outstanding-change', kpi.outstandingChange, false);

  // Advances
  document.getElementById('kpi-advances').textContent = formatCurrency(kpi.advances);
  updateChangeIndicator('kpi-advances-change', kpi.advancesChange, false);

  // Suspense
  document.getElementById('kpi-suspense').textContent = formatCurrency(kpi.suspense);
  updateChangeIndicator('kpi-suspense-change', kpi.suspenseChange, false);
}

function updateChangeIndicator(elementId, change, higherIsGood) {
  const element = document.getElementById(elementId);
  if (!element) return;

  const isPositive = change > 0;
  const isGood = higherIsGood ? isPositive : !isPositive;

  element.className = 'kpi-change ' + (change === 0 ? '' : (isGood ? 'positive' : 'negative'));
  element.innerHTML = `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <polyline points="${isPositive ? '23 6 13.5 15.5 8.5 10.5 1 18' : '23 18 13.5 8.5 8.5 13.5 1 6'}"/>
    </svg>
    ${change === 0 ? '0%' : (isPositive ? '+' : '') + change.toFixed(1) + '%'}
  `;
}

function updateYTD(ytd) {
  document.getElementById('ytd-received').textContent = formatCurrency(ytd.received);
  document.getElementById('ytd-payments').textContent = formatCurrency(ytd.payments);

  const netflowEl = document.getElementById('ytd-netflow');
  netflowEl.textContent = formatCurrency(ytd.netFlow);
  netflowEl.className = 'ytd-value ' + (ytd.netFlow >= 0 ? 'positive' : 'negative');

  document.getElementById('ytd-months').textContent = ytd.months;
}

/* ==================== Charts ==================== */
function updateCharts(data) {
  createCashFlowChart(data.banks);
  createOutstandingPieChart(data.outstanding.salesmen);
  createBankTrendChart(data.banks);
  createOutstandingTrendChart(data.outstanding);
  createSalesmanDistChart(data.outstanding.salesmen);
  createAdvancesTrendChart(data.advances);
  createSuspenseTrendChart(data.suspense);
}

function createCashFlowChart(data) {
  const ctx = document.getElementById('cashFlowChart');
  if (!ctx) return;

  if (charts.cashFlow) charts.cashFlow.destroy();

  charts.cashFlow = new Chart(ctx, {
    type: 'line',
    data: {
      labels: data.labels,
      datasets: [
        {
          label: 'Balance',
          data: data.balance,
          borderColor: '#4F46E5',
          backgroundColor: 'rgba(79, 70, 229, 0.1)',
          fill: true,
          tension: 0.4,
          borderWidth: 2
        },
        {
          label: 'Received',
          data: data.received,
          borderColor: '#10B981',
          backgroundColor: 'transparent',
          borderDash: [5, 5],
          tension: 0.4,
          borderWidth: 2
        },
        {
          label: 'Payments',
          data: data.payments,
          borderColor: '#EF4444',
          backgroundColor: 'transparent',
          borderDash: [5, 5],
          tension: 0.4,
          borderWidth: 2
        }
      ]
    },
    options: getLineChartOptions()
  });
}

function createOutstandingPieChart(salesmen) {
  const ctx = document.getElementById('outstandingPieChart');
  if (!ctx) return;

  if (charts.outstandingPie) charts.outstandingPie.destroy();

  const top5 = salesmen.slice(0, 5);
  const othersValue = salesmen.slice(5).reduce((sum, s) => sum + s.value, 0);

  const labels = top5.map(s => capitalizeFirst(s.name));
  const values = top5.map(s => s.value);

  if (othersValue > 0) {
    labels.push('Others');
    values.push(othersValue);
  }

  charts.outstandingPie = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: [
          '#4F46E5',
          '#10B981',
          '#F59E0B',
          '#EF4444',
          '#8B5CF6',
          '#6B7280'
        ],
        borderWidth: 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            usePointStyle: true,
            padding: 15,
            font: { size: 11 }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const total = context.dataset.data.reduce((a, b) => a + b, 0);
              const percentage = ((context.raw / total) * 100).toFixed(1);
              return context.label + ': ' + formatCurrency(context.raw) + ' (' + percentage + '%)';
            }
          }
        }
      },
      cutout: '60%'
    }
  });
}

function createBankTrendChart(data) {
  const ctx = document.getElementById('bankTrendChart');
  if (!ctx) return;

  if (charts.bankTrend) charts.bankTrend.destroy();

  charts.bankTrend = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: data.labels,
      datasets: [
        {
          label: 'Received',
          data: data.received,
          backgroundColor: 'rgba(16, 185, 129, 0.8)',
          borderRadius: 4
        },
        {
          label: 'Payments',
          data: data.payments,
          backgroundColor: 'rgba(239, 68, 68, 0.8)',
          borderRadius: 4
        }
      ]
    },
    options: getBarChartOptions()
  });
}

function createOutstandingTrendChart(data) {
  const ctx = document.getElementById('outstandingTrendChart');
  if (!ctx) return;

  if (charts.outstandingTrend) charts.outstandingTrend.destroy();

  charts.outstandingTrend = new Chart(ctx, {
    type: 'line',
    data: {
      labels: data.labels,
      datasets: [{
        label: 'Total Outstanding',
        data: data.total,
        borderColor: '#EF4444',
        backgroundColor: 'rgba(239, 68, 68, 0.1)',
        fill: true,
        tension: 0.4,
        borderWidth: 2
      }]
    },
    options: getSingleLineChartOptions()
  });
}

function createSalesmanDistChart(salesmen) {
  const ctx = document.getElementById('salesmanDistChart');
  if (!ctx) return;

  if (charts.salesmanDist) charts.salesmanDist.destroy();

  const top10 = salesmen.slice(0, 10);

  charts.salesmanDist = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: top10.map(s => capitalizeFirst(s.name)),
      datasets: [{
        label: 'Outstanding',
        data: top10.map(s => s.value),
        backgroundColor: top10.map((s, i) => {
          const colors = ['#4F46E5', '#10B981', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899', '#06B6D4', '#84CC16', '#F97316', '#6366F1'];
          return colors[i % colors.length];
        }),
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(context) {
              return formatCurrency(context.raw);
            }
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: {
            callback: function(value) {
              return formatCompact(value);
            }
          }
        },
        y: {
          grid: { display: false }
        }
      }
    }
  });
}

function createAdvancesTrendChart(data) {
  const ctx = document.getElementById('advancesTrendChart');
  if (!ctx) return;

  if (charts.advancesTrend) charts.advancesTrend.destroy();

  charts.advancesTrend = new Chart(ctx, {
    type: 'line',
    data: {
      labels: data.labels,
      datasets: [
        {
          label: 'Closing Balance',
          data: data.closing,
          borderColor: '#F59E0B',
          backgroundColor: 'rgba(245, 158, 11, 0.1)',
          fill: true,
          tension: 0.4,
          borderWidth: 2
        },
        {
          label: 'Given',
          data: data.given,
          borderColor: '#10B981',
          backgroundColor: 'transparent',
          borderDash: [5, 5],
          tension: 0.4,
          borderWidth: 2
        },
        {
          label: 'Settled',
          data: data.settled,
          borderColor: '#4F46E5',
          backgroundColor: 'transparent',
          borderDash: [5, 5],
          tension: 0.4,
          borderWidth: 2
        }
      ]
    },
    options: getLineChartOptions()
  });

  // Update advances KPIs
  const latestIdx = data.closing.length - 1;
  document.getElementById('adv-opening').textContent = formatCurrency(data.opening[latestIdx]);
  document.getElementById('adv-given').textContent = formatCurrency(data.given[latestIdx]);
  document.getElementById('adv-settled').textContent = formatCurrency(data.settled[latestIdx]);
  document.getElementById('adv-closing').textContent = formatCurrency(data.closing[latestIdx]);
}

function createSuspenseTrendChart(data) {
  const ctx = document.getElementById('suspenseTrendChart');
  if (!ctx) return;

  if (charts.suspenseTrend) charts.suspenseTrend.destroy();

  charts.suspenseTrend = new Chart(ctx, {
    type: 'line',
    data: {
      labels: data.labels,
      datasets: [{
        label: 'Suspense Balance',
        data: data.balance,
        borderColor: '#F59E0B',
        backgroundColor: 'rgba(245, 158, 11, 0.1)',
        fill: true,
        tension: 0.4,
        borderWidth: 2
      }]
    },
    options: getSingleLineChartOptions()
  });

  // Update suspense KPIs
  const latestIdx = data.balance.length - 1;
  document.getElementById('sus-opening').textContent = formatCurrency(data.balance[0]);
  document.getElementById('sus-debits').textContent = '--';
  document.getElementById('sus-credits').textContent = '--';
  document.getElementById('sus-closing').textContent = formatCurrency(data.balance[latestIdx]);
}

/* ==================== Chart Options ==================== */
function getLineChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: 'top',
        labels: {
          usePointStyle: true,
          padding: 20
        }
      },
      tooltip: {
        callbacks: {
          label: function(context) {
            return context.dataset.label + ': ' + formatCurrency(context.raw);
          }
        }
      }
    },
    scales: {
      y: {
        beginAtZero: false,
        ticks: {
          callback: function(value) {
            return formatCompact(value);
          }
        },
        grid: {
          color: 'rgba(0, 0, 0, 0.05)'
        }
      },
      x: {
        grid: { display: false }
      }
    },
    interaction: {
      intersect: false,
      mode: 'index'
    }
  };
}

function getSingleLineChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: {
        callbacks: {
          label: function(context) {
            return formatCurrency(context.raw);
          }
        }
      }
    },
    scales: {
      y: {
        beginAtZero: false,
        ticks: {
          callback: function(value) {
            return formatCompact(value);
          }
        }
      },
      x: {
        grid: { display: false }
      }
    }
  };
}

function getBarChartOptions() {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'top' },
      tooltip: {
        callbacks: {
          label: function(context) {
            return context.dataset.label + ': ' + formatCurrency(context.raw);
          }
        }
      }
    },
    scales: {
      y: {
        beginAtZero: true,
        ticks: {
          callback: function(value) {
            return formatCompact(value);
          }
        }
      },
      x: {
        grid: { display: false }
      }
    }
  };
}

/* ==================== Tables ==================== */
function updateTables(data) {
  updateTopOutstandingTable(data.outstanding.salesmen);
  updateBankSummaryTable(data.bankAccounts);
  updateOutstandingTable(data.outstanding.salesmen);
}

function updateTopOutstandingTable(salesmen) {
  const tbody = document.getElementById('top-outstanding-table');
  if (!tbody) return;

  const top5 = salesmen.slice(0, 5);
  tbody.innerHTML = top5.map(s => `
    <tr>
      <td>${capitalizeFirst(s.name)}</td>
      <td class="amount">${formatCurrency(s.value)}</td>
      <td class="${s.trend > 0 ? 'trend-up' : 'trend-down'}">
        ${s.trend > 0 ? '▲' : '▼'} ${Math.abs(s.trend).toFixed(1)}%
      </td>
    </tr>
  `).join('');
}

function updateBankSummaryTable(banks) {
  const tbody = document.getElementById('bank-summary-table');
  if (!tbody) return;

  tbody.innerHTML = banks.map(b => `
    <tr>
      <td>${b.name}</td>
      <td class="amount">${formatCurrency(b.balance)}</td>
      <td class="${b.change >= 0 ? 'positive' : 'negative'}">
        ${b.change >= 0 ? '+' : ''}${b.change.toFixed(1)}%
      </td>
    </tr>
  `).join('');
}

function updateOutstandingTable(salesmen) {
  const table = document.getElementById('outstanding-table');
  if (!table) return;

  table.innerHTML = `
    <thead>
      <tr>
        <th>Salesman</th>
        <th>Outstanding</th>
        <th>Trend</th>
      </tr>
    </thead>
    <tbody>
      ${salesmen.map(s => `
        <tr>
          <td>${capitalizeFirst(s.name)}</td>
          <td class="amount">${formatCurrency(s.value)}</td>
          <td class="${s.trend > 0 ? 'trend-up' : 'trend-down'}">
            ${s.trend > 0 ? '▲' : '▼'} ${Math.abs(s.trend).toFixed(1)}%
          </td>
        </tr>
      `).join('')}
    </tbody>
  `;
}

/* ==================== UI Functions ==================== */
function showSection(sectionId) {
  // Update navigation
  document.querySelectorAll('.nav-item').forEach(item => {
    item.classList.remove('active');
    if (item.getAttribute('data-section') === sectionId) {
      item.classList.add('active');
    }
  });

  // Update sections
  document.querySelectorAll('.section').forEach(section => {
    section.classList.remove('active');
  });

  const targetSection = document.getElementById(sectionId + '-section');
  if (targetSection) {
    targetSection.classList.add('active');
  }

  // Close mobile sidebar
  document.getElementById('sidebar').classList.remove('open');
}

function toggleSidebar() {
  const sidebar = document.getElementById('sidebar');
  sidebar.classList.toggle('collapsed');
  sidebar.classList.toggle('open');
}

function populateMonthSelector(months) {
  const select = document.getElementById('monthSelect');
  if (!select) return;

  select.innerHTML = months.slice().reverse().map(m =>
    `<option value="${m}" ${m === currentMonth ? 'selected' : ''}>${m}</option>`
  ).join('');
}

function changeMonth(month) {
  currentMonth = month;
  // In demo mode, we just update the display
  // With API, this would fetch new data
  showToast('info', 'Month Changed', `Viewing data for ${month}`);
}

function refreshData() {
  loadDashboardData();
  showToast('success', 'Refreshing', 'Data is being refreshed...');
}

function updateChartRange(range) {
  // Filter chart data based on range
  console.log('Chart range:', range);
  showToast('info', 'Range Updated', `Showing ${range === 'all' ? 'all' : range + ' months'} data`);
}

/* ==================== Actions ==================== */
function exportReport() {
  showToast('info', 'Exporting', 'Preparing PDF export...');

  // Use browser print for PDF export
  setTimeout(() => {
    window.print();
  }, 500);
}

function generateReport(type) {
  showToast('info', 'Report', `Generating ${type} report...`);

  // Switch to relevant section
  switch(type) {
    case 'monthly':
      showSection('overview');
      break;
    case 'comparison':
      showSection('banks');
      break;
    case 'outstanding':
      showSection('outstanding');
      break;
  }
}

/* ==================== Utilities ==================== */
function formatCurrency(value) {
  if (value === null || value === undefined || isNaN(value)) return '--';

  const currency = settings.currency || 'SAR';
  return new Intl.NumberFormat('en-SA', {
    style: 'currency',
    currency: currency,
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(value);
}

function formatCompact(value) {
  if (Math.abs(value) >= 1000000) {
    return (value / 1000000).toFixed(1) + 'M';
  } else if (Math.abs(value) >= 1000) {
    return (value / 1000).toFixed(0) + 'K';
  }
  return value.toFixed(0);
}

function capitalizeFirst(str) {
  if (!str) return '';
  return str.charAt(0).toUpperCase() + str.slice(1);
}

function showLoading(show) {
  const overlay = document.getElementById('loadingOverlay');
  if (overlay) {
    overlay.classList.toggle('active', show);
  }
}

function showToast(type, title, message) {
  const container = document.getElementById('toastContainer');
  if (!container) return;

  const icons = {
    success: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
    error: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
    warning: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
    info: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>'
  };

  const toast = document.createElement('div');
  toast.className = 'toast ' + type;
  toast.innerHTML = `
    <span class="toast-icon">${icons[type] || icons.info}</span>
    <div class="toast-content">
      <div class="toast-title">${title}</div>
      <div class="toast-message">${message}</div>
    </div>
    <button class="toast-close" onclick="this.parentElement.remove()">
      <svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2">
        <line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>
      </svg>
    </button>
  `;

  container.appendChild(toast);

  // Auto remove after 5 seconds
  setTimeout(() => {
    if (toast.parentElement) {
      toast.style.animation = 'slideIn 0.3s ease reverse';
      setTimeout(() => toast.remove(), 300);
    }
  }, 5000);
}
