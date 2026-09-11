const GAS_URL = "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec";

let currentUser = {
  citizenId: "",
  employeeId: "",
  role: "USER",
  tempPassword: "",
  payrolls: []
};

// Elements
const loginCard = document.getElementById("loginCard");
const changePwdCard = document.getElementById("changePwdCard");
const dashboardCard = document.getElementById("dashboardCard");
const adminModal = document.getElementById("adminModal");

// Đăng Nhập
document.getElementById("loginForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const citizenId = document.getElementById("citizenId").value.trim();
  const password = document.getElementById("password").value.trim();
  const msg = document.getElementById("loginMsg");
  msg.innerText = "Đang xác thực...";

  try {
    const res = await fetch(GAS_URL, {
      method: "POST",
      body: JSON.stringify({ action: "login", citizenId, password })
    }).then(r => r.json());

    if (!res.ok) {
      msg.innerText = res.message;
      return;
    }

    currentUser.citizenId = citizenId;
    currentUser.employeeId = res.employeeId;
    currentUser.role = res.role;
    currentUser.tempPassword = password;
    currentUser.payrolls = res.payrolls || [];

    msg.innerText = "";

    if (res.mustChangePassword) {
      loginCard.classList.add("hidden");
      changePwdCard.classList.remove("hidden");
    } else {
      showDashboard();
    }
  } catch (err) {
    msg.innerText = "Lỗi kết nối máy chủ!";
  }
});

// Đổi Mật Khẩu
document.getElementById("changePwdForm").addEventListener("submit", async (e) => {
  e.preventDefault();
  const newPassword = document.getElementById("newPassword").value.trim();
  const confirmPassword = document.getElementById("confirmPassword").value.trim();
  const msg = document.getElementById("pwdMsg");

  if (newPassword !== confirmPassword) {
    msg.innerText = "Mật khẩu xác nhận không khớp!";
    return;
  }

  msg.innerText = "Đang đổi mật khẩu...";

  try {
    const res = await fetch(GAS_URL, {
      method: "POST",
      body: JSON.stringify({
        action: "changePassword",
        citizenId: currentUser.citizenId,
        oldPassword: currentUser.tempPassword,
        newPassword: newPassword
      })
    }).then(r => r.json());

    if (!res.ok) {
      msg.innerText = res.message;
      return;
    }

    currentUser.payrolls = res.payrolls || [];
    changePwdCard.classList.add("hidden");
    showDashboard();
  } catch (err) {
    msg.innerText = "Lỗi đổi mật khẩu!";
  }
});

// Hiển Thị Bảng Lương Dashboard
function showDashboard() {
  loginCard.classList.add("hidden");
  dashboardCard.classList.remove("hidden");

  document.getElementById("userInfo").innerText = `Họ tên & Lương: ${currentUser.citizenId}`;
  document.getElementById("userSubInfo").innerText = `Mã NV: ${currentUser.employeeId} | Quyền: ${currentUser.role}`;

  if (currentUser.role === "ADMIN") {
    document.getElementById("btnAdmin").classList.remove("hidden");
  }

  const select = document.getElementById("selectPayMonth");
  select.innerHTML = "";

  if (currentUser.payrolls.length === 0) {
    document.getElementById("payrollDetail").innerHTML = "<p>Không tìm thấy dữ liệu lương.</p>";
    return;
  }

  currentUser.payrolls.forEach((p, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    opt.innerText = `Tháng ${p.PayMonth} (${p.Factory})`;
    select.appendChild(opt);
  });

  select.addEventListener("change", () => renderPayroll(select.value));
  renderPayroll(0);
}

// Render Chi Tiết Lương Theo Format Tiền VNĐ
function renderPayroll(index) {
  const p = currentUser.payrolls[index];
  if (!p) return;

  const fmt = (val) => typeof val === "number" ? val.toLocaleString("vi-VN") + " đ" : (val || 0) + " đ";
  const fmtNum = (val) => typeof val === "number" ? val.toFixed(2).replace(/\.00$/, '') : (val || 0);

  const html = `
    <div class="payroll-section">
      <h3>Thông Tin Chung</h3>
      <div class="payroll-grid">
        <div class="grid-item"><span>Họ tên:</span><span class="val">${p.FullName}</span></div>
        <div class="grid-item"><span>Phòng ban:</span><span class="val">${p.Department}</span></div>
        <div class="grid-item"><span>Bộ phận:</span><span class="val">${p.Section}</span></div>
        <div class="grid-item"><span>Chức vụ:</span><span class="val">${p.Position}</span></div>
        <div class="grid-item"><span>Cập nhật lúc:</span><span class="val">${p.UpdatedAt}</span></div>
      </div>
    </div>

    <div class="payroll-section">
      <h3>A. Các Khoản Thu Nhập (Tổng: ${fmt(p.TotalIncome)})</h3>
      <div class="payroll-grid">
        <div class="grid-item"><span>1. Lương tháng:</span><span class="val">${fmt(p.MonthlySalary)}</span></div>
        <div class="grid-item"><span>• Lương cơ bản:</span><span class="val">${fmt(p.BasicSalary)}</span></div>
        <div class="grid-item"><span>• Số ngày làm việc:</span><span class="val">${fmtNum(p.WorkingDays)}</span></div>
        <div class="grid-item"><span>• Số ngày lễ:</span><span class="val">${fmtNum(p.HolidayDays)}</span></div>
        <div class="grid-item"><span>• Số ngày nghỉ hưởng lương:</span><span class="val">${fmtNum(p.PaidLeaveDays)}</span></div>
        <div class="grid-item"><span>• Số ngày nghỉ không lương:</span><span class="val">${fmtNum(p.UnpaidLeaveDays)}</span></div>
        <div class="grid-item"><span>• Số ngày nghỉ TL tối thiểu vùng:</span><span class="val">${fmtNum(p.RegionalMinimumLeaveDays)}</span></div>

        <div class="grid-item"><span>2. Lương tăng ca / Đêm:</span><span class="val">${fmt(p.OvertimePay)}</span></div>
        <div class="grid-item"><span>• Giờ làm ngoài giờ:</span><span class="val">${fmtNum(p.OvertimeHours)}</span></div>
        <div class="grid-item"><span>• Giờ làm ngày nghỉ:</span><span class="val">${fmtNum(p.RestDayHours)}</span></div>
        <div class="grid-item"><span>• Giờ ngoài giờ ngày nghỉ:</span><span class="val">${fmtNum(p.OvertimeRestDayHours)}</span></div>
        <div class="grid-item"><span>• Tăng ca đêm ngày nghỉ:</span><span class="val">${fmtNum(p.NightRestDayOvertimeHours)}</span></div>
        <div class="grid-item"><span>• Giờ làm ngày lễ:</span><span class="val">${fmtNum(p.HolidayHours)}</span></div>
        <div class="grid-item"><span>• Giờ ngoài giờ ngày lễ:</span><span class="val">${fmtNum(p.HolidayOvertimeHours)}</span></div>
        <div class="grid-item"><span>• Tăng ca đêm ngày lễ:</span><span class="val">${fmtNum(p.NightHolidayOvertimeHours)}</span></div>
        <div class="grid-item"><span>• Số ngày làm ca đêm:</span><span class="val">${fmtNum(p.NightShiftDays)}</span></div>
        <div class="grid-item"><span>• Giờ tăng ca đêm:</span><span class="val">${fmtNum(p.NightOvertimeHours)}</span></div>

        <div class="grid-item"><span>3. Phụ cấp, hỗ trợ khác:</span><span class="val"></span></div>
        <div class="grid-item"><span>• Cơm / Công tác phí:</span><span class="val">${fmt(p.OtherMoney)}</span></div>
        <div class="grid-item"><span>• Tiền kỷ luật:</span><span class="val">${fmt(p.Discipline)}</span></div>
        <div class="grid-item"><span>• Tiền gắn bó 2 năm:</span><span class="val">${fmt(p.Loyalty2Years)}</span></div>
        <div class="grid-item"><span>• Tiền gắn bó 5 năm:</span><span class="val">${fmt(p.Loyalty5Years)}</span></div>
        <div class="grid-item"><span>• Tiền gắn bó 10 năm:</span><span class="val">${fmt(p.Loyalty10Years)}</span></div>
        <div class="grid-item"><span>• Tiền nhà ở:</span><span class="val">${fmt(p.Housing)}</span></div>
        <div class="grid-item"><span>• Tiền đi lại:</span><span class="val">${fmt(p.Transportation)}</span></div>
        <div class="grid-item"><span>• Thưởng chuyên cần:</span><span class="val">${fmt(p.AttendanceBonus)}</span></div>
        <div class="grid-item"><span>• Hoa hồng & Thưởng vượt định mức:</span><span class="val">${fmt(p.SalesCommissionBonus)}</span></div>
        <div class="grid-item"><span>• Trợ cấp thôi việc & Phép tồn:</span><span class="val">${fmt(p.SeveranceUnusedLeave)}</span></div>
      </div>
    </div>

    <div class="payroll-section">
      <h3>B. Các Khoản Khấu Trừ</h3>
      <div class="payroll-grid">
        <div class="grid-item"><span>BHXH:</span><span class="val">${fmt(p.SocialInsurance)}</span></div>
        <div class="grid-item"><span>BHYT:</span><span class="val">${fmt(p.HealthInsurance)}</span></div>
        <div class="grid-item"><span>BHTN:</span><span class="val">${fmt(p.UnemploymentInsurance)}</span></div>
        <div class="grid-item"><span>Thuế TNCN:</span><span class="val">${fmt(p.PersonalIncomeTax)}</span></div>
        <div class="grid-item"><span>Tạm ứng:</span><span class="val">${fmt(p.Advance)}</span></div>
        <div class="grid-item"><span>Khấu trừ khác:</span><span class="val">${fmt(p.OtherDeductions)}</span></div>
      </div>
    </div>

    <div class="payroll-section">
      <div class="payroll-grid">
        <div class="grid-item highlight"><span>C. LƯƠNG THỰC LĨNH:</span><span class="val">${fmt(p.NetPay)}</span></div>
      </div>
    </div>
  `;

  document.getElementById("payrollDetail").innerHTML = html;
}

// Chức năng Admin Modal
document.getElementById("btnAdmin").addEventListener("click", () => adminModal.classList.remove("hidden"));
document.getElementById("btnCloseAdmin").addEventListener("click", () => adminModal.classList.add("hidden"));

document.getElementById("btnConfirmReset").addEventListener("click", async () => {
  const targetCitizenId = document.getElementById("targetCitizenId").value.trim();
  const msg = document.getElementById("adminMsg");
  if (!targetCitizenId) {
    msg.innerText = "Vui lòng nhập CCCD!";
    return;
  }
  msg.innerText = "Đang xử lý...";

  try {
    const res = await fetch(GAS_URL, {
      method: "POST",
      body: JSON.stringify({
        action: "adminResetPassword",
        adminCitizenId: currentUser.citizenId,
        targetCitizenId: targetCitizenId
      })
    }).then(r => r.json());

    msg.innerText = res.message;
  } catch (err) {
    msg.innerText = "Lỗi khi reset!";
  }
});

// Đăng Xuất
document.getElementById("btnLogout").addEventListener("click", () => {
  location.reload();
});
