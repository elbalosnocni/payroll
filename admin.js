/**
 * admin.js
 * -----------------------------------------------------------------------
 * Toàn bộ hành động ở đây gọi các action adminXxx trên GAS, backend sẽ tự
 * kiểm tra session.role === 'admin' trước khi xử lý (frontend chỉ ẩn/hiện
 * giao diện, không phải lớp bảo mật thật sự).
 * -----------------------------------------------------------------------
 */

document.addEventListener('DOMContentLoaded', () => {
  const adminNavBtn = document.getElementById('admin-nav-btn');
  const payrollNavBtn = document.getElementById('payroll-nav-btn');
  const payrollView = document.getElementById('payroll-view');
  const adminView = document.getElementById('admin-view');

  if (adminNavBtn) {
    adminNavBtn.addEventListener('click', () => {
      payrollView.style.display = 'none';
      adminView.style.display = 'block';
      switchAdminTab('list');
    });
  }
  if (payrollNavBtn) {
    payrollNavBtn.addEventListener('click', () => {
      adminView.style.display = 'none';
      payrollView.style.display = 'flex';
    });
  }

  document.querySelectorAll('.admin-tabs button').forEach(btn => {
    btn.addEventListener('click', () => switchAdminTab(btn.dataset.tab));
  });

  const searchForm = document.getElementById('admin-search-form');
  if (searchForm) {
    searchForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const keyword = document.getElementById('admin-search-input').value.trim();
      if (!keyword) return;
      const res = await callApi('adminSearchEmployee', { keyword });
      renderEmployeeTable(document.getElementById('admin-search-results'), res.ok ? res.data.results : [], res.ok ? null : res.error);
    });
  }
});

function switchAdminTab(tab) {
  document.querySelectorAll('.admin-tabs button').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
  document.querySelectorAll('.admin-panel').forEach(p => p.style.display = 'none');
  const panel = document.getElementById('admin-panel-' + tab);
  if (panel) panel.style.display = 'block';

  if (tab === 'list') loadEmployeeList();
  if (tab === 'sync') loadSyncStatus();
  if (tab === 'audit') loadAuditLog();
}

async function loadEmployeeList() {
  const container = document.getElementById('admin-employee-list');
  container.innerHTML = '<div class="empty-state">Đang tải...</div>';
  const res = await callApi('adminListEmployees', {});
  renderEmployeeTable(container, res.ok ? res.data.employees : [], res.ok ? null : res.error);
}

function renderEmployeeTable(container, employees, error) {
  if (error) {
    container.innerHTML = `<div class="empty-state">${escapeHtml(error)}</div>`;
    return;
  }
  if (!employees || !employees.length) {
    container.innerHTML = '<div class="empty-state">Không có dữ liệu.</div>';
    return;
  }

  const rows = employees.map(e => `
    <tr>
      <td>${escapeHtml(e.maNV)}</td>
      <td>${escapeHtml(e.hoTen)}</td>
      <td>${escapeHtml(e.cccd)}</td>
      <td>${escapeHtml(e.xuong || '')}</td>
      <td>${escapeHtml(e.chucVu || '')}</td>
      <td>${e.role === 'admin' ? '<span class="badge">Admin</span>' : ''}</td>
      <td>${e.mustChangePassword ? '<span class="badge">Chờ đổi MK</span>' : ''}</td>
      <td><button class="action-link" data-manv="${escapeHtml(e.maNV)}" data-name="${escapeHtml(e.hoTen)}" onclick="confirmResetPassword(this)">Reset mật khẩu</button></td>
    </tr>
  `).join('');

  container.innerHTML = `
    <table class="data-table">
      <thead><tr>
        <th>Mã NV</th><th>Họ tên</th><th>CCCD</th><th>Xưởng</th><th>Chức vụ</th><th>Vai trò</th><th>Trạng thái</th><th></th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

async function confirmResetPassword(btn) {
  const maNV = btn.dataset.manv;
  const name = btn.dataset.name;
  if (!confirm(`Đặt lại mật khẩu của "${name}" về Mã nhân viên (${maNV})?`)) return;

  btn.disabled = true;
  btn.textContent = 'Đang xử lý...';
  const res = await callApi('adminResetPassword', { maNV });
  if (res.ok) {
    alert(res.data.message);
    loadEmployeeList();
  } else {
    alert(res.error || 'Reset mật khẩu thất bại.');
    btn.disabled = false;
    btn.textContent = 'Reset mật khẩu';
  }
}

async function loadSyncStatus() {
  const container = document.getElementById('admin-sync-status');
  container.innerHTML = '<div class="empty-state">Đang tải...</div>';
  const res = await callApi('adminSyncStatus', {});
  if (!res.ok) {
    container.innerHTML = `<div class="empty-state">${escapeHtml(res.error)}</div>`;
    return;
  }
  const rows = (res.data.syncStatus || []).map(s => `
    <tr>
      <td>${escapeHtml(s.Xuong)}</td>
      <td>${escapeHtml(s.Thang)}</td>
      <td>${escapeHtml(s.LastSyncAt)}</td>
      <td>${escapeHtml(s.SoDongDaXuLy)}</td>
      <td>${escapeHtml(s.TrangThai)}</td>
      <td>${escapeHtml(s.GhiChu || '')}</td>
    </tr>
  `).join('');
  container.innerHTML = rows ? `
    <table class="data-table">
      <thead><tr><th>Xưởng</th><th>Tháng</th><th>Đồng bộ lúc</th><th>Số dòng</th><th>Trạng thái</th><th>Ghi chú</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>` : '<div class="empty-state">Chưa có dữ liệu đồng bộ.</div>';
}

async function loadAuditLog() {
  const container = document.getElementById('admin-audit-log');
  container.innerHTML = '<div class="empty-state">Đang tải...</div>';
  const res = await callApi('adminAuditLog', { limit: 200 });
  if (!res.ok) {
    container.innerHTML = `<div class="empty-state">${escapeHtml(res.error)}</div>`;
    return;
  }
  const rows = (res.data.logs || []).map(l => `
    <tr>
      <td>${escapeHtml(l.Timestamp)}</td>
      <td>${escapeHtml(l.Actor)}</td>
      <td>${escapeHtml(l.Action)}</td>
      <td>${escapeHtml(l.Target)}</td>
      <td>${escapeHtml(l.Detail || '')}</td>
    </tr>
  `).join('');
  container.innerHTML = rows ? `
    <table class="data-table">
      <thead><tr><th>Thời gian</th><th>Người thực hiện</th><th>Hành động</th><th>Đối tượng</th><th>Chi tiết</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>` : '<div class="empty-state">Chưa có log nào.</div>';
}
