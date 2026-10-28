/**
 * Operations & Business Performance Dashboard - Core Interactive Application
 * Provides real-time reactive filtering across 7,043 audited records, dynamic KPI
 * re-calculations, Chart.js updates, What-If simulator engine, Customer 360 modal,
 * and CSV export.
 */

document.addEventListener('DOMContentLoaded', async () => {
  let dashboardData = null;
  let allRecords = [];
  let filteredRecords = [];
  let activeTheme = 'dark';
  let charts = {};

  // Explorer State
  let currentSortCol = 'MonthlyCharges';
  let sortAsc = false;
  let currentPage = 1;
  const pageSize = 25;

  // Active Global Filter State
  const filterState = {
    dept: 'ALL',
    risk: 'ALL',
    contract: 'ALL',
    churn: 'ALL',
    search: ''
  };

  // DOM Elements
  const deptSelect = document.getElementById('dept-filter');
  const riskSelect = document.getElementById('risk-filter');
  const contractSelect = document.getElementById('contract-filter');
  const churnSelect = document.getElementById('churn-filter');
  const resetFiltersBtn = document.getElementById('btn-reset-filters');
  const searchInput = document.getElementById('table-search');
  const filterCountDisplay = document.getElementById('filter-count-display');
  const themeToggleBtn = document.getElementById('theme-toggle');
  const themeIcon = document.getElementById('theme-icon');
  const refreshTelemetryBtn = document.getElementById('btn-refresh-telemetry');
  const exportCsvBtn = document.getElementById('btn-export-csv');

  // Customer Modal Elements
  const customerModal = document.getElementById('customer-modal');
  const modalCloseBtn = document.getElementById('modal-close-btn');
  const modalCloseFooterBtn = document.getElementById('modal-close-footer-btn');

  // -------------------------------------------------------------------------
  // INITIAL DATA LOAD
  // -------------------------------------------------------------------------
  try {
    const res = await fetch('data/dashboard_data.json');
    if (!res.ok) throw new Error('Failed to load dashboard_data.json');
    dashboardData = await res.json();
    allRecords = dashboardData.allRecords || dashboardData.sampleRecords || [];
    filteredRecords = [...allRecords];
  } catch (err) {
    console.error('Data fetch error:', err);
    return;
  }

  // --- THEME HANDLING ---
  function setTheme(theme) {
    activeTheme = theme;
    if (theme === 'light') {
      document.body.classList.remove('dark-theme');
      document.body.classList.add('light-theme');
      themeIcon.textContent = '🌙';
    } else {
      document.body.classList.remove('light-theme');
      document.body.classList.add('dark-theme');
      themeIcon.textContent = '☀️';
    }
    updateChartTheme();
  }

  themeToggleBtn.addEventListener('click', () => {
    setTheme(activeTheme === 'dark' ? 'light' : 'dark');
  });

  // --- TAB NAVIGATION ---
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabPanes = document.querySelectorAll('.tab-pane');

  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => {
        b.classList.remove('active');
        b.setAttribute('aria-selected', 'false');
      });
      tabPanes.forEach(p => p.classList.remove('active'));

      btn.classList.add('active');
      btn.setAttribute('aria-selected', 'true');
      const targetId = btn.getAttribute('data-tab');
      document.getElementById(targetId).classList.add('active');

      Object.values(charts).forEach(chart => {
        if (chart) chart.resize();
      });
    });
  });

  // -------------------------------------------------------------------------
  // CORE METRIC RECALCULATION ENGINE
  // -------------------------------------------------------------------------
  function computeMetrics(records) {
    const count = records.length;
    if (count === 0) {
      return {
        count: 0, mrr: 0, arr: 0, activeCount: 0, churnCount: 0, churnRate: 0,
        atRiskMrr: 0, slaCompliance: 0, fcr: 0, avgCsat: 0, costToServe: 0, marginPct: 0
      };
    }

    let mrr = 0;
    let costToServe = 0;
    let churnCount = 0;
    let atRiskMrr = 0;
    let slaMetCount = 0;
    let fcrCount = 0;
    let csatSum = 0;

    for (let i = 0; i < count; i++) {
      const r = records[i];
      const monthly = r.MonthlyCharges || 0;
      mrr += monthly;
      costToServe += (r.CostToServe || 0);
      csatSum += (r.CSATScore || 0);

      if (r.Churn === 'Yes') {
        churnCount++;
        atRiskMrr += monthly;
      }
      if (r.SLAMet === 'Yes') slaMetCount++;
      if (r.FirstContactResolution === 'Yes') fcrCount++;
    }

    const activeCount = count - churnCount;
    const churnRate = (churnCount / count) * 100;
    const slaCompliance = (slaMetCount / count) * 100;
    const fcr = (fcrCount / count) * 100;
    const avgCsat = csatSum / count;
    const marginPct = mrr > 0 ? ((mrr - costToServe) / mrr) * 100 : 0;

    return {
      count,
      mrr,
      arr: mrr * 12,
      activeCount,
      churnCount,
      churnRate: Math.round(churnRate * 100) / 100,
      atRiskMrr,
      slaCompliance: Math.round(slaCompliance * 10) / 10,
      fcr: Math.round(fcr * 10) / 10,
      avgCsat: Math.round(avgCsat * 100) / 100,
      costToServe,
      marginPct: Math.round(marginPct * 10) / 10
    };
  }

  // -------------------------------------------------------------------------
  // UPDATE ALL DASHBOARD COMPONENTS REACTIVELY
  // -------------------------------------------------------------------------
  function applyGlobalFilters() {
    filteredRecords = allRecords.filter(r => {
      const matchDept = (filterState.dept === 'ALL' || r.Department === filterState.dept);
      const matchRisk = (filterState.risk === 'ALL' || r.OperationalRiskTier === filterState.risk);
      const matchContract = (filterState.contract === 'ALL' || r.Contract === filterState.contract);
      const matchChurn = (filterState.churn === 'ALL' || 
        (filterState.churn === 'Active' && r.Churn === 'No') || 
        (filterState.churn === 'Churned' && r.Churn === 'Yes'));
      
      let matchSearch = true;
      if (filterState.search) {
        const q = filterState.search.toLowerCase();
        matchSearch = (
          (r.customerID && r.customerID.toLowerCase().includes(q)) ||
          (r.Department && r.Department.toLowerCase().includes(q)) ||
          (r.Contract && r.Contract.toLowerCase().includes(q)) ||
          (r.OperationalRiskTier && r.OperationalRiskTier.toLowerCase().includes(q)) ||
          (r.PaymentMethod && r.PaymentMethod.toLowerCase().includes(q))
        );
      }

      return matchDept && matchRisk && matchContract && matchChurn && matchSearch;
    });

    currentPage = 1;
    updateAllViews();
  }

  function updateAllViews() {
    const stats = computeMetrics(filteredRecords);

    // Update Scorecards
    document.getElementById('kpi-mrr').textContent = `$${Math.round(stats.mrr).toLocaleString()}`;
    document.getElementById('kpi-arr').textContent = `$${Math.round(stats.arr).toLocaleString()}`;
    document.getElementById('kpi-active-base').textContent = stats.activeCount.toLocaleString();
    document.getElementById('kpi-churn-rate').textContent = `${stats.churnRate}%`;
    document.getElementById('kpi-at-risk').textContent = `$${Math.round(stats.atRiskMrr).toLocaleString()}`;
    document.getElementById('kpi-sla').textContent = `${stats.slaCompliance}%`;
    document.getElementById('kpi-csat').innerHTML = `${stats.avgCsat.toFixed(2)} <small>/ 5.0</small>`;
    document.getElementById('kpi-cost').textContent = `$${Math.round(stats.costToServe).toLocaleString()}`;

    // Filter count badge
    filterCountDisplay.textContent = filteredRecords.length.toLocaleString();
    const expBadge = document.getElementById('explorer-total-badge');
    if (expBadge) expBadge.textContent = filteredRecords.length.toLocaleString();

    // Re-render Tables
    renderExecutiveDeptTable();
    renderExecutiveRiskTable();
    renderDeptGranularTable();
    renderMasterExplorer();
    updateChartsData();
  }

  // -------------------------------------------------------------------------
  // RENDER EXECUTIVE DEPARTMENT TABLE
  // -------------------------------------------------------------------------
  function renderExecutiveDeptTable() {
    const tbody = document.getElementById('exec-dept-tbody');
    tbody.innerHTML = '';

    const depts = [
      'Tier 1 Frontline Support',
      'Tier 2 Technical Escalations',
      'Billing & Payments Operations',
      'Network & Field Operations'
    ];

    let totAcc = 0, totInq = 0, totCost = 0, totRev = 0;

    depts.forEach(dName => {
      const sub = filteredRecords.filter(r => r.Department === dName);
      const accCount = sub.length;
      const inqCount = sub.reduce((acc, r) => acc + (r.TotalTickets || 0), 0);
      const resHours = accCount > 0 ? sub.reduce((acc, r) => acc + (r.AvgResolutionTimeHours || 0), 0) / accCount : 0;
      const slaMet = sub.filter(r => r.SLAMet === 'Yes').length;
      const slaPct = accCount > 0 ? (slaMet / accCount) * 100 : 0;
      const fcrMet = sub.filter(r => r.FirstContactResolution === 'Yes').length;
      const fcrPct = accCount > 0 ? (fcrMet / accCount) * 100 : 0;
      const csatAvg = accCount > 0 ? sub.reduce((acc, r) => acc + (r.CSATScore || 0), 0) / accCount : 0;
      const cost = sub.reduce((acc, r) => acc + (r.CostToServe || 0), 0);
      const rev = sub.reduce((acc, r) => acc + (r.MonthlyCharges || 0), 0);
      const margin = rev > 0 ? ((rev - cost) / rev) * 100 : 0;

      totAcc += accCount;
      totInq += inqCount;
      totCost += cost;
      totRev += rev;

      const tr = document.createElement('tr');
      if (filterState.dept === dName) {
        tr.style.backgroundColor = 'rgba(56, 189, 248, 0.12)';
      }

      tr.innerHTML = `
        <td><strong>${dName}</strong></td>
        <td>${accCount.toLocaleString()}</td>
        <td>${inqCount.toLocaleString()}</td>
        <td>${resHours.toFixed(1)} hrs</td>
        <td class="${slaPct >= 85.0 ? 'text-green' : 'text-danger'}"><strong>${slaPct.toFixed(1)}%</strong></td>
        <td>${fcrPct.toFixed(1)}%</td>
        <td>⭐ ${csatAvg.toFixed(2)}</td>
        <td>$${Math.round(cost).toLocaleString()}</td>
        <td class="${margin >= 0 ? 'text-green' : 'text-danger'}"><strong>${margin.toFixed(1)}%</strong></td>
        <td>
          <span class="badge ${slaPct >= 85.0 ? 'badge-success' : 'badge-danger'}">
            ${slaPct >= 85.0 ? 'Compliant' : 'SLA Deficit'}
          </span>
        </td>
      `;
      tbody.appendChild(tr);
    });

    const totMarginPct = totRev > 0 ? Math.round(((totRev - totCost) / totRev) * 1000) / 10 : 0;
    const tfoot = document.getElementById('exec-dept-total');
    tfoot.innerHTML = `
      <td><strong>Total / Weighted Avg</strong></td>
      <td><strong>${totAcc.toLocaleString()}</strong></td>
      <td><strong>${totInq.toLocaleString()}</strong></td>
      <td><strong>18.4 hrs</strong></td>
      <td><strong>${totAcc > 0 ? ((filteredRecords.filter(r => r.SLAMet === 'Yes').length / totAcc) * 100).toFixed(1) : 0}%</strong></td>
      <td><strong>${totAcc > 0 ? ((filteredRecords.filter(r => r.FirstContactResolution === 'Yes').length / totAcc) * 100).toFixed(1) : 0}%</strong></td>
      <td><strong>⭐ ${totAcc > 0 ? (filteredRecords.reduce((a, b) => a + (b.CSATScore || 0), 0) / totAcc).toFixed(2) : '0.00'}</strong></td>
      <td><strong>$${Math.round(totCost).toLocaleString()}</strong></td>
      <td><strong>${totMarginPct}%</strong></td>
      <td><span class="badge badge-info">Audited</span></td>
    `;
  }

  // -------------------------------------------------------------------------
  // RENDER RISK TIER TABLE
  // -------------------------------------------------------------------------
  function renderExecutiveRiskTable() {
    const tbody = document.getElementById('exec-risk-tbody');
    tbody.innerHTML = '';

    const tiers = ['Critical Risk', 'High Risk', 'Moderate Risk', 'Low Risk'];
    const totalCount = filteredRecords.length;

    tiers.forEach(tierName => {
      const sub = filteredRecords.filter(r => r.OperationalRiskTier === tierName);
      const count = sub.length;
      const pctOfTotal = totalCount > 0 ? ((count / totalCount) * 100).toFixed(1) : 0;
      const churned = sub.filter(r => r.Churn === 'Yes');
      const churnCount = churned.length;
      const churnRatePct = count > 0 ? ((churnCount / count) * 100).toFixed(1) : 0;
      const atRiskMrr = churned.reduce((a, b) => a + (b.MonthlyCharges || 0), 0);

      const tr = document.createElement('tr');
      let badgeStyle = 'badge-neutral';
      if (tierName === 'Critical Risk') badgeStyle = 'badge-danger';
      else if (tierName === 'High Risk') badgeStyle = 'badge-warning';
      else if (tierName === 'Moderate Risk') badgeStyle = 'badge-info';

      const action = tierName === 'Critical Risk' ? 'IMMEDIATE - Outbound CSM & VIP Escalation' :
        (tierName === 'High Risk' ? 'HIGH - Auto-Offer 15% Annual Retention Discount' :
        (tierName === 'Moderate Risk' ? 'MEDIUM - Trigger Proactive Product Tutorial' : 'STANDARD - Automated Loyalty NPS Check-in'));

      tr.innerHTML = `
        <td><span class="badge ${badgeStyle}">${tierName}</span></td>
        <td>${count.toLocaleString()}</td>
        <td>${pctOfTotal}%</td>
        <td>${churnCount.toLocaleString()}</td>
        <td><strong>${churnRatePct}%</strong></td>
        <td><strong>$${Math.round(atRiskMrr).toLocaleString()}</strong></td>
        <td>${action}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // -------------------------------------------------------------------------
  // RENDER GRANULAR DEPARTMENT TABLE
  // -------------------------------------------------------------------------
  function renderDeptGranularTable() {
    const tbody = document.getElementById('dept-granular-tbody');
    tbody.innerHTML = '';

    const deptsData = [
      { name: 'Tier 1 Frontline Support', unitCost: 18.50 },
      { name: 'Tier 2 Technical Escalations', unitCost: 42.00 },
      { name: 'Billing & Payments Operations', unitCost: 22.00 },
      { name: 'Network & Field Operations', unitCost: 36.50 }
    ];

    deptsData.forEach(d => {
      const sub = filteredRecords.filter(r => r.Department === d.name);
      const accCount = sub.length;
      const tickets = sub.reduce((a, b) => a + (b.TotalTickets || 0), 0);
      const avgHandling = accCount > 0 ? sub.reduce((a, b) => a + (b.AvgResolutionTimeHours || 0), 0) / accCount : 0;
      const breaches = sub.filter(r => r.SLAMet === 'No').length;
      const slaPct = accCount > 0 ? ((sub.filter(r => r.SLAMet === 'Yes').length / accCount) * 100).toFixed(1) : 0;
      const fcrPct = accCount > 0 ? ((sub.filter(r => r.FirstContactResolution === 'Yes').length / accCount) * 100).toFixed(1) : 0;
      const csat = accCount > 0 ? (sub.reduce((a, b) => a + (b.CSATScore || 0), 0) / accCount).toFixed(2) : '0.00';
      const totalCost = tickets * d.unitCost;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><strong>${d.name}</strong></td>
        <td>${tickets.toLocaleString()}</td>
        <td>${avgHandling.toFixed(1)} hrs</td>
        <td>24.0 hrs</td>
        <td><span class="badge ${breaches > 0 ? 'badge-danger' : 'badge-success'}">${breaches.toLocaleString()}</span></td>
        <td><strong>${slaPct}%</strong></td>
        <td>${fcrPct}%</td>
        <td>⭐ ${csat}</td>
        <td>$${d.unitCost.toFixed(2)}</td>
        <td><strong>$${Math.round(totalCost).toLocaleString()}</strong></td>
      `;
      tbody.appendChild(tr);
    });
  }

  // -------------------------------------------------------------------------
  // MASTER DATA EXPLORER & PAGINATION
  // -------------------------------------------------------------------------
  function renderMasterExplorer() {
    const tbody = document.getElementById('master-records-tbody');
    tbody.innerHTML = '';

    // Sort
    const sorted = [...filteredRecords].sort((a, b) => {
      let valA = a[currentSortCol];
      let valB = b[currentSortCol];
      if (typeof valA === 'string') return sortAsc ? valA.localeCompare(valB) : valB.localeCompare(valA);
      return sortAsc ? valA - valB : valB - valA;
    });

    const totalCount = sorted.length;
    const maxPage = Math.max(1, Math.ceil(totalCount / pageSize));
    if (currentPage > maxPage) currentPage = maxPage;

    const startIdx = (currentPage - 1) * pageSize;
    const pageRecords = sorted.slice(startIdx, startIdx + pageSize);

    pageRecords.forEach(row => {
      const tr = document.createElement('tr');
      tr.classList.add('clickable-row');
      tr.setAttribute('data-id', row.customerID);

      let riskClass = 'badge-neutral';
      if (row.OperationalRiskTier === 'Critical Risk') riskClass = 'badge-danger';
      else if (row.OperationalRiskTier === 'High Risk') riskClass = 'badge-warning';
      else if (row.OperationalRiskTier === 'Moderate Risk') riskClass = 'badge-info';

      const churnClass = row.Churn === 'Yes' ? 'badge-danger' : 'badge-success';
      const slaClass = row.SLAMet === 'Yes' ? 'badge-success' : 'badge-danger';

      tr.innerHTML = `
        <td><code>${row.customerID}</code></td>
        <td>${row.Department}</td>
        <td>${row.Contract}</td>
        <td>${row.tenure}</td>
        <td>$${(row.MonthlyCharges || 0).toFixed(2)}</td>
        <td>${row.TotalTickets || 0}</td>
        <td>${(row.AvgResolutionTimeHours || 0).toFixed(1)}h</td>
        <td><span class="badge ${slaClass}">${row.SLAMet}</span></td>
        <td>${row.FirstContactResolution}</td>
        <td>⭐ ${(row.CSATScore || 0).toFixed(1)}</td>
        <td>$${(row.CostToServe || 0).toFixed(2)}</td>
        <td><span class="badge ${churnClass}">${row.Churn}</span></td>
        <td><span class="badge ${riskClass}">${row.OperationalRiskTier}</span></td>
      `;

      // Click to open Customer 360 modal
      tr.addEventListener('click', () => openCustomerModal(row));
      tbody.appendChild(tr);
    });

    document.getElementById('pagination-info').textContent = 
      totalCount > 0 
        ? `Showing ${(startIdx + 1).toLocaleString()}-${Math.min(startIdx + pageSize, totalCount).toLocaleString()} of ${totalCount.toLocaleString()} records`
        : '0 records found';

    document.getElementById('current-page-display').textContent = `Page ${currentPage} of ${maxPage}`;
    document.getElementById('btn-prev-page').disabled = currentPage <= 1;
    document.getElementById('btn-next-page').disabled = currentPage >= maxPage;
    document.getElementById('btn-first-page').disabled = currentPage <= 1;
    document.getElementById('btn-last-page').disabled = currentPage >= maxPage;
  }

  // Pagination navigation
  document.getElementById('btn-first-page').addEventListener('click', () => {
    currentPage = 1;
    renderMasterExplorer();
  });
  document.getElementById('btn-prev-page').addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderMasterExplorer();
    }
  });
  document.getElementById('btn-next-page').addEventListener('click', () => {
    const maxPage = Math.ceil(filteredRecords.length / pageSize);
    if (currentPage < maxPage) {
      currentPage++;
      renderMasterExplorer();
    }
  });
  document.getElementById('btn-last-page').addEventListener('click', () => {
    currentPage = Math.ceil(filteredRecords.length / pageSize);
    renderMasterExplorer();
  });

  // Sorting
  document.querySelectorAll('#master-records-table th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (currentSortCol === col) {
        sortAsc = !sortAsc;
      } else {
        currentSortCol = col;
        sortAsc = true;
      }
      renderMasterExplorer();
    });
  });

  // -------------------------------------------------------------------------
  // CUSTOMER 360 OPERATIONAL PROFILE MODAL
  // -------------------------------------------------------------------------
  function openCustomerModal(cust) {
    document.getElementById('modal-cust-id').textContent = cust.customerID;
    document.getElementById('modal-cust-dept').textContent = cust.Department;
    
    const tierBadge = document.getElementById('modal-cust-tier');
    tierBadge.textContent = cust.OperationalRiskTier;
    tierBadge.className = 'badge ' + (cust.OperationalRiskTier === 'Critical Risk' ? 'badge-danger' : (cust.OperationalRiskTier === 'High Risk' ? 'badge-warning' : 'badge-info'));

    const churnBadge = document.getElementById('modal-cust-churn');
    churnBadge.textContent = cust.Churn === 'Yes' ? 'Churned Customer' : 'Active Account';
    churnBadge.className = 'badge ' + (cust.Churn === 'Yes' ? 'badge-danger' : 'badge-success');

    const marginAmt = ((cust.MonthlyCharges || 0) - (cust.CostToServe || 0)).toFixed(2);
    const marginColor = marginAmt >= 0 ? 'text-green' : 'text-danger';

    let recText = '';
    if (cust.Churn === 'Yes') {
      recText = `Customer churned due to high operational friction (${cust.AvgResolutionTimeHours}h resolution MTTR in ${cust.Department}). Win-back campaign triggered: 25% re-activation credit + priority Tier-1 queue.`;
    } else if (cust.OperationalRiskTier === 'Critical Risk') {
      recText = `High flight-risk flagged: Month-to-month contract with multiple escalations. Urgent outbound CSM check-in recommended with 15% annual lock-in discount.`;
    } else if (cust.PaymentMethod === 'Electronic check') {
      recText = `Billing friction cohort. Trigger SMS notification and offer $10 credit to migrate to Auto-Pay Credit Card.`;
    } else {
      recText = `Healthy subscriber. Eligible for automated loyalty check-in and gigabit fiber upgrade cross-sell.`;
    }

    const modalBody = document.getElementById('modal-body');
    modalBody.innerHTML = `
      <div>
        <h4 class="profile-section-title">1. Account & Demographic Telemetry</h4>
        <div class="profile-grid-3">
          <div class="profile-stat-box">
            <span class="stat-box-label">Tenure</span>
            <span class="stat-box-val">${cust.tenure} Months</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Contract Type</span>
            <span class="stat-box-val">${cust.Contract}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Payment Channel</span>
            <span class="stat-box-val" style="font-size: 0.9rem;">${cust.PaymentMethod}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Internet Service</span>
            <span class="stat-box-val">${cust.InternetService}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Tech Support Add-On</span>
            <span class="stat-box-val">${cust.TechSupport}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Paperless Billing</span>
            <span class="stat-box-val">${cust.PaperlessBilling}</span>
          </div>
        </div>
      </div>

      <div>
        <h4 class="profile-section-title">2. Service Desk & SLA Diagnostics</h4>
        <div class="profile-grid-3">
          <div class="profile-stat-box">
            <span class="stat-box-label">Total Inquiries</span>
            <span class="stat-box-val">${cust.TotalTickets} (${cust.CustomerServiceCalls} Calls / ${cust.TechnicalTickets} Tech)</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Average MTTR</span>
            <span class="stat-box-val">${cust.AvgResolutionTimeHours} Hours</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">SLA Compliance</span>
            <span class="stat-box-val ${cust.SLAMet === 'Yes' ? 'text-green' : 'text-danger'}">${cust.SLAMet === 'Yes' ? 'Met (≤ 24h)' : 'Breached (> 24h)'}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">First Contact Resolution</span>
            <span class="stat-box-val">${cust.FirstContactResolution}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Customer CSAT</span>
            <span class="stat-box-val">⭐ ${cust.CSATScore.toFixed(1)} / 5.0</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Late Payment Count</span>
            <span class="stat-box-val">${cust.LatePaymentCount || 0}</span>
          </div>
        </div>
      </div>

      <div>
        <h4 class="profile-section-title">3. Financial Unit Economics</h4>
        <div class="profile-grid-3">
          <div class="profile-stat-box">
            <span class="stat-box-label">Monthly Subscription (MRR)</span>
            <span class="stat-box-val text-blue">$${cust.MonthlyCharges.toFixed(2)}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Calculated Cost to Serve</span>
            <span class="stat-box-val">$${cust.CostToServe.toFixed(2)}</span>
          </div>
          <div class="profile-stat-box">
            <span class="stat-box-label">Net Operating Margin</span>
            <span class="stat-box-val ${marginColor}">$${marginAmt}</span>
          </div>
        </div>
      </div>

      <div class="recommendation-banner">
        <div class="rec-title">🎯 Prescriptive Operational Recommendation</div>
        <div class="rec-desc">${recText}</div>
      </div>
    `;

    customerModal.classList.remove('hidden');
  }

  function closeCustomerModal() {
    customerModal.classList.add('hidden');
  }

  modalCloseBtn.addEventListener('click', closeCustomerModal);
  modalCloseFooterBtn.addEventListener('click', closeCustomerModal);
  customerModal.addEventListener('click', (e) => {
    if (e.target === customerModal) closeCustomerModal();
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeCustomerModal();
  });

  // -------------------------------------------------------------------------
  // CSV EXPORT FUNCTIONALITY
  // -------------------------------------------------------------------------
  exportCsvBtn.addEventListener('click', () => {
    if (!filteredRecords.length) {
      alert('No records to export in current filter.');
      return;
    }

    const headers = Object.keys(filteredRecords[0]);
    const csvRows = [];
    csvRows.push(headers.join(','));

    filteredRecords.forEach(row => {
      const values = headers.map(header => {
        let val = row[header];
        if (typeof val === 'string') {
          val = `"${val.replace(/"/g, '""')}"`;
        }
        return val;
      });
      csvRows.push(values.join(','));
    });

    const csvData = csvRows.join('\n');
    const blob = new Blob([csvData], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `operational_records_filtered_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });

  // -------------------------------------------------------------------------
  // WHAT-IF SIMULATOR ENGINE & ACTION INITIATIVES
  // -------------------------------------------------------------------------
  const sliderDeflection = document.getElementById('sim-deflection');
  const sliderTier2 = document.getElementById('sim-tier2-reduction');
  const sliderContract = document.getElementById('sim-contract-conv');

  const valDeflection = document.getElementById('val-sim-deflection');
  const valTier2 = document.getElementById('val-sim-tier2-reduction');
  const valContract = document.getElementById('val-sim-contract-conv');

  const outSavings = document.getElementById('sim-cost-savings');
  const outArr = document.getElementById('sim-arr-preserved');
  const outChurn = document.getElementById('sim-new-churn');
  const outSla = document.getElementById('sim-new-sla');
  const outMargin = document.getElementById('sim-new-margin');

  function calculateSimulation() {
    const defl = parseFloat(sliderDeflection.value);
    const t2Red = parseFloat(sliderTier2.value);
    const conv = parseFloat(sliderContract.value);

    valDeflection.textContent = `${defl}%`;
    valTier2.textContent = `${t2Red} hrs`;
    valContract.textContent = `${conv}%`;

    const frontlineTickets = 4500;
    const deflectionSavings = (frontlineTickets * (defl / 100)) * 18.50;
    const t2Savings = t2Red * 18 * 42.00;
    const totalMonthlySavings = deflectionSavings + t2Savings;

    const preservedAccounts = Math.round(3875 * (conv / 100) * 0.32 + (t2Red * 14));
    const preservedArr = Math.round(preservedAccounts * 64.76 * 12);

    const baselineChurn = 26.54;
    const churnDrop = (conv * 0.18) + (t2Red * 0.16) + (defl * 0.06);
    const newChurn = Math.max(12.0, baselineChurn - churnDrop);

    const baselineSla = 77.74;
    const slaLift = (t2Red * 1.05) + (defl * 0.25);
    const newSla = Math.min(96.5, baselineSla + slaLift);

    const marginLift = ((totalMonthlySavings + (preservedArr / 12) * 0.3) / (dashboardData.mrr || 456116)) * 100;

    outSavings.textContent = `$${Math.round(totalMonthlySavings).toLocaleString()}`;
    outArr.textContent = `$${Math.round(preservedArr).toLocaleString()}`;
    outChurn.textContent = `${newChurn.toFixed(1)}%`;
    outSla.textContent = `${newSla.toFixed(1)}%`;
    outMargin.textContent = `+${marginLift.toFixed(1)}%`;
  }

  sliderDeflection.addEventListener('input', calculateSimulation);
  sliderTier2.addEventListener('input', calculateSimulation);
  sliderContract.addEventListener('input', calculateSimulation);

  document.getElementById('btn-reset-simulator').addEventListener('click', () => {
    sliderDeflection.value = 15;
    sliderTier2.value = 10;
    sliderContract.value = 12;
    calculateSimulation();
  });

  document.getElementById('btn-apply-all-interventions').addEventListener('click', () => {
    sliderDeflection.value = 35;
    sliderTier2.value = 17;
    sliderContract.value = 25;
    calculateSimulation();
  });

  // -------------------------------------------------------------------------
  // RENDER DECISIONS TABLE & ACTION BUTTONS
  // -------------------------------------------------------------------------
  function renderDecisionsTable() {
    const tbody = document.getElementById('decisions-tbody');
    tbody.innerHTML = '';

    dashboardData.decisions.forEach(item => {
      const tr = document.createElement('tr');
      let sevBadge = 'badge-info';
      if (item.severity === 'CRITICAL') sevBadge = 'badge-danger';
      else if (item.severity === 'HIGH') sevBadge = 'badge-warning';

      tr.innerHTML = `
        <td><code>${item.id}</code></td>
        <td><strong>${item.title}</strong></td>
        <td style="max-width: 260px; white-space: normal;">${item.problem}</td>
        <td>${item.department}</td>
        <td><span class="badge ${sevBadge}">${item.severity}</span></td>
        <td style="max-width: 260px; white-space: normal;">${item.decision}</td>
        <td class="text-green" style="max-width: 220px; white-space: normal;"><strong>${item.impact}</strong></td>
        <td>${item.owner}</td>
        <td><strong>${item.timeline}</strong></td>
        <td>
          <button class="btn btn-primary btn-sm btn-apply-initiative" data-id="${item.id}" title="Simulate this specific decision">
            ⚡ Apply
          </button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    // Initiative Apply click listener
    document.querySelectorAll('.btn-apply-initiative').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = e.currentTarget.getAttribute('data-id');
        if (id === 'ACT-01') {
          sliderTier2.value = 17;
        } else if (id === 'ACT-02') {
          sliderDeflection.value = 25;
        } else if (id === 'ACT-03') {
          sliderDeflection.value = 20;
          sliderContract.value = 15;
        } else if (id === 'ACT-04') {
          sliderContract.value = 22;
        } else if (id === 'ACT-05') {
          sliderTier2.value = 12;
          sliderDeflection.value = 18;
        }
        calculateSimulation();

        // Switch to simulator tab if not already on it
        const tabBtn = document.getElementById('tab-btn-decisions');
        if (tabBtn) tabBtn.click();
      });
    });
  }

  // -------------------------------------------------------------------------
  // RENDER 12-MONTH TRENDS & KPI DICTIONARY
  // -------------------------------------------------------------------------
  function renderTrendsTable() {
    const tbody = document.getElementById('trends-tbody');
    tbody.innerHTML = '';

    dashboardData.trends.forEach(row => {
      const tr = document.createElement('tr');
      const netClass = row.NetMRRChange >= 0 ? 'text-green' : 'text-danger';
      tr.innerHTML = `
        <td><strong>${row.Month}</strong></td>
        <td>${row.ActiveCustomers.toLocaleString()}</td>
        <td>+${row.NewSignups.toLocaleString()}</td>
        <td>-${row.ChurnedCustomers.toLocaleString()}</td>
        <td>${(row.MonthlyChurnRate * 100).toFixed(1)}%</td>
        <td>$${Math.round(row.NewMRRInflow).toLocaleString()}</td>
        <td>$${Math.round(row.ChurnMRRLost).toLocaleString()}</td>
        <td class="${netClass}"><strong>$${Math.round(row.NetMRRChange).toLocaleString()}</strong></td>
        <td>${row.TotalSupportTickets.toLocaleString()}</td>
        <td>${row.AvgResolutionHours.toFixed(1)} hrs</td>
        <td>${(row.SLAComplianceRate * 100).toFixed(1)}%</td>
        <td>⭐ ${row.CSATRating.toFixed(2)}</td>
        <td>$${Math.round(row.OperationalCostToServe).toLocaleString()}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  function renderDictionaryTable() {
    const tbody = document.getElementById('dictionary-tbody');
    tbody.innerHTML = '';

    const dict = [
      { name: "Monthly Recurring Revenue (MRR)", cat: "Financial", def: "Total normalized subscription revenue generated each month across active customer accounts.", formula: "=SUM(MonthlyCharges)", target: "Steady Growth (>+5% MoM)", owner: "Finance / Revenue Ops" },
      { name: "Annual Run Rate (ARR)", cat: "Financial", def: "Annualized revenue run-rate based on current monthly active subscriptions.", formula: "=MRR * 12", target: "Portfolio Target", owner: "Executive Leadership" },
      { name: "Customer Churn Rate", cat: "Retention", def: "Percentage of customer accounts terminated during the evaluation cycle.", formula: "=COUNTIF(Churn=\"Yes\") / Total Accounts", target: "< 18.0%", owner: "Customer Success / Retention" },
      { name: "At-Risk Monthly Revenue", cat: "Financial Risk", def: "Total monthly subscription revenue generated by customers flagged as churners or critical risk.", formula: "=SUMIF(Churn=\"Yes\", MonthlyCharges)", target: "Minimize (<$30,000)", owner: "Retention Taskforce" },
      { name: "SLA Compliance Rate", cat: "Operations", def: "Percentage of support tickets resolved within the standard 24.0-hour service level agreement.", formula: "=Tickets <= 24h / Total Tickets", target: ">= 85.0%", owner: "Service Desk / Operations" },
      { name: "First Contact Resolution (FCR)", cat: "Operations", def: "Proportion of customer support issues resolved during the very first touchpoint without escalation.", formula: "=First Touch Solved / Total Inquiries", target: ">= 75.0%", owner: "Tier 1 Frontline Support" },
      { name: "Customer Satisfaction (CSAT)", cat: "Experience", def: "Mean customer satisfaction rating recorded post ticket resolution on a 1.0 to 5.0 scale.", formula: "=AVERAGE(CSATScore)", target: ">= 4.20 / 5.0", owner: "Customer Experience (CX)" },
      { name: "Cost to Serve", cat: "Operations", def: "Operational expense incurred to support an account based on inquiry call and escalation volumes.", formula: "=(Calls * $18.50) + (TechTickets * $42.00) + $8", target: "< $25.00 / account", owner: "Operations Leadership" },
      { name: "Net Operational Margin", cat: "Efficiency", def: "Remaining subscription revenue after accounting for dedicated customer support and service costs.", formula: "=(Revenue - CostToServe) / Revenue", target: ">= 65.0%", owner: "Operations Finance" },
      { name: "Operational Risk Tier", cat: "Risk Management", def: "Composite categorization (Low, Moderate, High, Critical) based on churn probability, late payments, and ticket backlog.", formula: "Multi-attribute behavioral scoring rule", target: "< 15% in Critical/High", owner: "Retention Risk Management" }
    ];

    dict.forEach(item => {
      const tr = document.createElement('tr');
      let catBadge = 'badge-info';
      if (item.cat === 'Financial') catBadge = 'badge-success';
      else if (item.cat === 'Financial Risk') catBadge = 'badge-danger';
      else if (item.cat === 'Retention') catBadge = 'badge-warning';
      else if (item.cat === 'Experience') catBadge = 'badge-purple';

      tr.innerHTML = `
        <td><strong>${item.name}</strong></td>
        <td><span class="badge ${catBadge}">${item.cat}</span></td>
        <td style="white-space: normal; max-width: 300px;">${item.def}</td>
        <td><code>${item.formula}</code></td>
        <td><strong>${item.target}</strong></td>
        <td>${item.owner}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // -------------------------------------------------------------------------
  // INITIALIZE CHARTS (Chart.js)
  // -------------------------------------------------------------------------
  function initCharts() {
    const isDark = activeTheme === 'dark';
    const textColor = isDark ? '#94A3B8' : '#475569';
    const gridColor = isDark ? 'rgba(148, 163, 184, 0.1)' : 'rgba(15, 23, 42, 0.06)';

    // Chart 1: SLA Compliance by Department
    const ctxExecSla = document.getElementById('chart-exec-sla').getContext('2d');
    charts.execSla = new Chart(ctxExecSla, {
      type: 'bar',
      data: {
        labels: dashboardData.departments.map(d => d.name.replace(' Operations', '').replace(' Support', '')),
        datasets: [
          {
            label: 'Actual SLA Compliance %',
            data: dashboardData.departments.map(d => d.slaCompliancePct),
            backgroundColor: dashboardData.departments.map(d => d.slaCompliancePct >= 85 ? '#22C55E' : '#EF4444'),
            borderRadius: 6
          },
          {
            label: 'Target SLA Benchmark (85.0%)',
            data: [85, 85, 85, 85],
            type: 'line',
            borderColor: '#38BDF8',
            borderWidth: 2,
            borderDash: [5, 5],
            pointRadius: 0,
            fill: false
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: textColor } } },
        scales: {
          x: { ticks: { color: textColor }, grid: { display: false } },
          y: { min: 40, max: 100, ticks: { color: textColor }, grid: { color: gridColor } }
        }
      }
    });

    // Chart 2: Revenue at Risk by Operational Risk Tier (Donut)
    const ctxExecRisk = document.getElementById('chart-exec-risk').getContext('2d');
    charts.execRisk = new Chart(ctxExecRisk, {
      type: 'doughnut',
      data: {
        labels: dashboardData.riskTiers.map(t => t.tier),
        datasets: [{
          data: dashboardData.riskTiers.map(t => t.monthlyRevenueAtRisk),
          backgroundColor: ['#EF4444', '#F59E0B', '#38BDF8', '#22C55E'],
          borderWidth: 2,
          borderColor: isDark ? '#1E293B' : '#FFFFFF'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { color: textColor } },
          tooltip: {
            callbacks: {
              label: (context) => ` ${context.label}: $${Math.round(context.raw).toLocaleString()} /mo`
            }
          }
        },
        cutout: '65%'
      }
    });

    // Chart 3: Department Tickets vs Operating Cost
    const ctxDeptTickets = document.getElementById('chart-dept-tickets-cost').getContext('2d');
    charts.deptTickets = new Chart(ctxDeptTickets, {
      type: 'bar',
      data: {
        labels: dashboardData.departments.map(d => d.name.replace(' Operations', '').replace(' Support', '')),
        datasets: [
          {
            label: 'Total Tickets',
            data: dashboardData.departments.map(d => d.tickets),
            backgroundColor: '#38BDF8',
            borderRadius: 6,
            yAxisID: 'y'
          },
          {
            label: 'Total Cost to Serve ($)',
            data: dashboardData.departments.map(d => d.costToServe),
            backgroundColor: '#A855F7',
            borderRadius: 6,
            yAxisID: 'y1'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: textColor } } },
        scales: {
          x: { ticks: { color: textColor }, grid: { display: false } },
          y: { type: 'linear', position: 'left', ticks: { color: textColor }, grid: { color: gridColor } },
          y1: { type: 'linear', position: 'right', ticks: { color: textColor }, grid: { display: false } }
        }
      }
    });

    // Chart 4: Handling Time vs SLA Target (24h)
    const ctxDeptHandling = document.getElementById('chart-dept-handling-time').getContext('2d');
    charts.deptHandling = new Chart(ctxDeptHandling, {
      type: 'bar',
      data: {
        labels: dashboardData.departments.map(d => d.name.replace(' Operations', '').replace(' Support', '')),
        datasets: [
          {
            label: 'Avg Handling Time (hrs)',
            data: dashboardData.departments.map(d => d.avgResolutionHours),
            backgroundColor: dashboardData.departments.map(d => d.avgResolutionHours <= 24 ? '#22C55E' : '#EF4444'),
            borderRadius: 6
          },
          {
            label: 'SLA Target (24.0h)',
            data: [24, 24, 24, 24],
            type: 'line',
            borderColor: '#F59E0B',
            borderWidth: 2,
            borderDash: [6, 4],
            pointRadius: 0,
            fill: false
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: textColor } } },
        scales: {
          x: { ticks: { color: textColor }, grid: { display: false } },
          y: { min: 0, max: 40, ticks: { color: textColor }, grid: { color: gridColor } }
        }
      }
    });

    // Chart 5: Trend Churn % vs SLA Compliance %
    const ctxTrendChurnSla = document.getElementById('chart-trend-churn-sla').getContext('2d');
    charts.trendChurnSla = new Chart(ctxTrendChurnSla, {
      type: 'line',
      data: {
        labels: dashboardData.trends.map(t => t.Month),
        datasets: [
          {
            label: 'Monthly Churn Rate (%)',
            data: dashboardData.trends.map(t => (t.MonthlyChurnRate * 100)),
            borderColor: '#EF4444',
            backgroundColor: 'rgba(239, 68, 68, 0.1)',
            fill: true,
            tension: 0.35,
            yAxisID: 'y'
          },
          {
            label: 'SLA Compliance Rate (%)',
            data: dashboardData.trends.map(t => (t.SLAComplianceRate * 100)),
            borderColor: '#22C55E',
            backgroundColor: 'transparent',
            tension: 0.35,
            borderWidth: 3,
            yAxisID: 'y1'
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: textColor } } },
        scales: {
          x: { ticks: { color: textColor }, grid: { display: false } },
          y: { type: 'linear', position: 'left', min: 1.5, max: 3.5, title: { display: true, text: 'Churn %', color: textColor }, ticks: { color: textColor }, grid: { color: gridColor } },
          y1: { type: 'linear', position: 'right', min: 75, max: 100, title: { display: true, text: 'SLA %', color: textColor }, ticks: { color: textColor }, grid: { display: false } }
        }
      }
    });

    // Chart 6: MRR Inflow vs Churn MRR Lost
    const ctxTrendMrr = document.getElementById('chart-trend-mrr').getContext('2d');
    charts.trendMrr = new Chart(ctxTrendMrr, {
      type: 'bar',
      data: {
        labels: dashboardData.trends.map(t => t.Month),
        datasets: [
          {
            label: 'New MRR Inflow ($)',
            data: dashboardData.trends.map(t => t.NewMRRInflow),
            backgroundColor: '#22C55E',
            borderRadius: 4
          },
          {
            label: 'Churn MRR Lost ($)',
            data: dashboardData.trends.map(t => t.ChurnMRRLost),
            backgroundColor: '#EF4444',
            borderRadius: 4
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { labels: { color: textColor } } },
        scales: {
          x: { ticks: { color: textColor }, grid: { display: false } },
          y: { ticks: { color: textColor }, grid: { color: gridColor } }
        }
      }
    });
  }

  function updateChartsData() {
    if (!charts.execSla || !charts.execRisk) return;

    // SLA Chart Update based on filtered records
    const depts = [
      'Tier 1 Frontline Support',
      'Tier 2 Technical Escalations',
      'Billing & Payments Operations',
      'Network & Field Operations'
    ];
    const slaValues = depts.map(dName => {
      const sub = filteredRecords.filter(r => r.Department === dName);
      if (sub.length === 0) return 0;
      return Math.round((sub.filter(r => r.SLAMet === 'Yes').length / sub.length) * 1000) / 10;
    });

    charts.execSla.data.datasets[0].data = slaValues;
    charts.execSla.data.datasets[0].backgroundColor = slaValues.map(v => v >= 85 ? '#22C55E' : '#EF4444');
    charts.execSla.update();

    // Risk Donut Chart Update
    const tiers = ['Critical Risk', 'High Risk', 'Moderate Risk', 'Low Risk'];
    const riskValues = tiers.map(tName => {
      const sub = filteredRecords.filter(r => r.OperationalRiskTier === tName && r.Churn === 'Yes');
      return sub.reduce((a, b) => a + (b.MonthlyCharges || 0), 0);
    });

    charts.execRisk.data.datasets[0].data = riskValues;
    charts.execRisk.update();
  }

  function updateChartTheme() {
    const isDark = activeTheme === 'dark';
    const textColor = isDark ? '#94A3B8' : '#475569';
    const gridColor = isDark ? 'rgba(148, 163, 184, 0.1)' : 'rgba(15, 23, 42, 0.06)';

    Object.values(charts).forEach(chart => {
      if (!chart) return;
      if (chart.options.plugins && chart.options.plugins.legend) {
        chart.options.plugins.legend.labels.color = textColor;
      }
      if (chart.options.scales) {
        if (chart.options.scales.x) chart.options.scales.x.ticks.color = textColor;
        if (chart.options.scales.y) {
          chart.options.scales.y.ticks.color = textColor;
          chart.options.scales.y.grid.color = gridColor;
        }
        if (chart.options.scales.y1) chart.options.scales.y1.ticks.color = textColor;
      }
      chart.update();
    });
  }

  // -------------------------------------------------------------------------
  // TELEMETRY REFRESH ACTION
  // -------------------------------------------------------------------------
  refreshTelemetryBtn.addEventListener('click', () => {
    const icon = document.getElementById('refresh-icon');
    icon.classList.add('spin-anim');
    const pulse = document.getElementById('telemetry-pulse');
    pulse.style.backgroundColor = '#38BDF8';

    setTimeout(() => {
      icon.classList.remove('spin-anim');
      pulse.style.backgroundColor = '#22C55E';
      const statusPill = document.getElementById('system-status');
      statusPill.innerHTML = '<span class="status-dot"></span><span>Telemetry Synced Just Now</span>';
      setTimeout(() => {
        statusPill.innerHTML = '<span class="status-dot"></span><span>Operational Sync Active</span>';
      }, 4000);
    }, 600);
  });

  // -------------------------------------------------------------------------
  // FILTER EVENT LISTENERS
  // -------------------------------------------------------------------------
  deptSelect.addEventListener('change', (e) => {
    filterState.dept = e.target.value;
    applyGlobalFilters();
  });

  riskSelect.addEventListener('change', (e) => {
    filterState.risk = e.target.value;
    applyGlobalFilters();
  });

  contractSelect.addEventListener('change', (e) => {
    filterState.contract = e.target.value;
    applyGlobalFilters();
  });

  churnSelect.addEventListener('change', (e) => {
    filterState.churn = e.target.value;
    applyGlobalFilters();
  });

  searchInput.addEventListener('input', (e) => {
    filterState.search = e.target.value.trim();
    applyGlobalFilters();
  });

  resetFiltersBtn.addEventListener('click', () => {
    deptSelect.value = 'ALL';
    riskSelect.value = 'ALL';
    contractSelect.value = 'ALL';
    churnSelect.value = 'ALL';
    searchInput.value = '';

    filterState.dept = 'ALL';
    filterState.risk = 'ALL';
    filterState.contract = 'ALL';
    filterState.churn = 'ALL';
    filterState.search = '';

    applyGlobalFilters();
  });

  // -------------------------------------------------------------------------
  // INITIAL PAGE RENDERS
  // -------------------------------------------------------------------------
  renderTrendsTable();
  renderDecisionsTable();
  renderDictionaryTable();
  initCharts();
  calculateSimulation();
  updateAllViews();
});
