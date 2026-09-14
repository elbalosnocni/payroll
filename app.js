/**
 * app.js
 * -----------------------------------------------------------------------
 * Hiển thị danh sách các tháng có phiếu lương và nội dung phiếu lương của
 * nhân viên đang đăng nhập. Server luôn xác định nhân viên qua token phiên,
 * frontend không bao giờ gửi MaNV/CCCD của người khác lên để lấy dữ liệu.
 * -----------------------------------------------------------------------
 */

let currentMonths = [];
let currentMonthSelected = null;

async function initPayrollView() {
  const monthList = document.getElementById('month-list');
  const main = document.getElementById('payslip-container');
  monthList.innerHTML = '<li><span class="empty-state">Đang tải...</span></li>';

  const res = await callApi('getPayroll', {});
  if (!res.ok) {
    monthList.innerHTML = '';
    main.innerHTML = `<div class="empty-state">${escapeHtml(res.error || 'Không tải được dữ liệu.')}</div>`;
    if (res.code === 'SESSION_EXPIRED') {
      Session.clear();
      setTimeout(() => location.reload(), 1500);
    }
    return;
  }

  document.getElementById('sidebar-name').textContent = res.data.hoTen || '';
  document.getElementById('sidebar-meta').textContent =
    [res.data.chucVu, res.data.boPhan, res.data.phongBan].filter(Boolean).join(' · ');

  currentMonths = res.data.months || [];

  if (!currentMonths.length) {
    monthList.innerHTML = '<li><span class="empty-state">Chưa có phiếu lương</span></li>';
    main.innerHTML = '<div class="empty-state">Chưa có dữ liệu lương nào được đồng bộ cho tài khoản này.</div>';
    return;
  }

  renderMonthList();
  await loadPayslip(currentMonths[0]);
}

function renderMonthList() {
  const monthList = document.getElementById('month-list');
  monthList.innerHTML = '';
  currentMonths.forEach(m => {
    const li = document.createElement('li');
    const btn = document.createElement('button');
    btn.textContent = formatMonthLabel(m);
    btn.className = m === currentMonthSelected ? 'active' : '';
    btn.addEventListener('click', () => loadPayslip(m));
    li.appendChild(btn);
    monthList.appendChild(li);
  });
}

function formatMonthLabel(thang) {
  // thang dạng "08-2026" -> "Tháng 08/2026"
  const parts = thang.split('-');
  if (parts.length === 2) return `Tháng ${parts[0]}/${parts[1]}`;
  return thang;
}

async function loadPayslip(thang) {
  currentMonthSelected = thang;
  renderMonthList();

  const main = document.getElementById('payslip-container');
  main.innerHTML = '<div class="empty-state">Đang tải phiếu lương...</div>';

  const res = await callApi('getPayroll', { thang });
  if (!res.ok) {
    main.innerHTML = `<div class="empty-state">${escapeHtml(res.error || 'Không tải được phiếu lương.')}</div>`;
    return;
  }

  main.innerHTML = renderPayslipHtml(res.data);
}

function renderPayslipHtml(d) {
  const info = d.thongTinNhanVien;
  const lt = d.luongThang;
  const ngg = d.luongNgoaiGio;
  const pc = d.phuCapTroCap;
  const kt = d.khauTru;

  const line = (label, value) => `
    <tr><td class="label">${escapeHtml(label)}</td><td class="value">${escapeHtml(value)}</td></tr>`;
  const subLine = (label, value) => `
    <tr class="sub"><td class="label">${escapeHtml(label)}</td><td class="value">${escapeHtml(value)}</td></tr>`;

  return `
    <div class="payslip">
      <div class="payslip-header">
        <h1>Phiếu lương</h1>
        <div class="period">${escapeHtml(formatMonthLabel(d.thang))}</div>
      </div>
      <dl class="payslip-meta">
        <div><dt>Họ tên:</dt><dd>${escapeHtml(info.hoTen)}</dd></div>
        <div><dt>Mã NV:</dt><dd>${escapeHtml(info.maNV)}</dd></div>
        <div><dt>Xưởng:</dt><dd>${escapeHtml(info.xuong || '')}</dd></div>
        <div><dt>Chức vụ:</dt><dd>${escapeHtml(info.chucVu || '')}</dd></div>
        <div><dt>Phòng ban:</dt><dd>${escapeHtml(info.phongBan || '')}</dd></div>
        <div><dt>Bộ phận:</dt><dd>${escapeHtml(info.boPhan || '')}</dd></div>
      </dl>

      <div class="section-title">A. Lương tháng</div>
      <table class="line-table">
        ${subLine('Lương cơ bản', lt.luongCoBan)}
        ${subLine('Số ngày làm việc', lt.soNgayLamViec)}
        ${subLine('Số ngày lễ', lt.soNgayLe)}
        ${subLine('Số ngày nghỉ hưởng lương', lt.soNgayNghiHuongLuong)}
        ${subLine('Số ngày nghỉ không hưởng lương', lt.soNgayNghiKhongLuong)}
        ${subLine('Số ngày nghỉ hưởng lương tối thiểu vùng', lt.soNgayNghiHuongLuongToiThieuVung)}
        ${line('Cộng lương tháng', lt.tongLuongThang)}
      </table>

      <div class="section-title">Lương ngoài giờ, ngày nghỉ, ca đêm</div>
      <table class="line-table">
        ${subLine('Số giờ ngoài giờ', ngg.soGioNgoaiGio)}
        ${subLine('Số giờ làm ngày nghỉ', ngg.soGioNgayNghi)}
        ${subLine('Số giờ ngoài giờ ngày nghỉ', ngg.soGioNgoaiGioNgayNghi)}
        ${subLine('Số giờ tăng ca đêm', ngg.soGioTangCaDem)}
        ${subLine('Số giờ làm ngày lễ', ngg.soGioLamNgayLe)}
        ${subLine('Số giờ ngoài giờ ngày lễ', ngg.soGioNgoaiGioNgayLe)}
        ${subLine('Số giờ tăng ca đêm ngày lễ', ngg.soGioTangCaDemNgayLe)}
        ${subLine('Số ngày làm ca đêm', ngg.soNgayLamCaDem)}
        ${line('Cộng lương ngoài giờ', ngg.tongLuongNgoaiGio)}
      </table>

      <div class="section-title">Phụ cấp, trợ cấp, hỗ trợ khác</div>
      <table class="line-table">
        ${line('Tiền khác (công tác phí, cơm)', pc.tienKhac)}
        ${line('Tiền kỷ luật', pc.tienKyLuat)}
        ${line('Tiền gắn bó 2 năm', pc.tienGanBo2Nam)}
        ${line('Tiền gắn bó 5 năm', pc.tienGanBo5Nam)}
        ${line('Tiền gắn bó 10 năm', pc.tienGanBo10Nam)}
        ${line('Tiền nhà ở', pc.tienNhaO)}
        ${line('Tiền đi lại', pc.tienDiLai)}
        ${line('Tiền thưởng chuyên cần', pc.tienThuongChuyenCan)}
        ${line('Hoa hồng & thưởng vượt định mức', pc.hoaHongThuongVuotDinhMuc)}
        ${line('Trợ cấp thôi việc & phép năm còn lại', pc.troCapThoiViecPhepNam)}
      </table>

      <table class="line-table">
        <tr class="total-row"><td class="label">Tổng khoản thu nhập</td><td class="value">${escapeHtml(d.tongThuNhap)}</td></tr>
      </table>

      <div class="section-title">B. Các khoản khấu trừ</div>
      <table class="line-table">
        ${line('BHXH', kt.bhxh)}
        ${line('BHYT', kt.bhyt)}
        ${line('BHTN', kt.bhtn)}
        ${line('Thuế thu nhập', kt.thueThuNhap)}
        ${line('Tạm ứng', kt.tamUng)}
        ${line('Khấu trừ khác', kt.khauTruKhac)}
      </table>

      <table class="line-table">
        <tr class="total-row grand"><td class="label">Lương thực lĩnh</td><td class="value">${escapeHtml(d.luongThucLinh)}</td></tr>
      </table>

      <div class="updated-note">Dữ liệu cập nhật lúc: ${escapeHtml(d.updatedAt || '')}</div>
    </div>
  `;
}

function escapeHtml(str) {
  return String(str == null ? '' : str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
