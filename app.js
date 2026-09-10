const API_URL =
  "https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec";

const state = {
  token: "",
  user: null,
  adminToken: ""
};

document.addEventListener("DOMContentLoaded", () => {
  buildMonthList();
  bindEvents();
});

const $ = id => document.getElementById(id);

function bindEvents() {
  $("loginBtn").addEventListener("click", login);
  $("changeBtn").addEventListener("click", changePassword);
  $("loadBtn").addEventListener("click", loadPayslip);
  $("logoutBtn").addEventListener("click", logout);
  $("adminLoginBtn").addEventListener("click", adminLogin);
  $("resetBtn").addEventListener("click", adminReset);

  ["username","password"].forEach(id =>
    $(id).addEventListener("keydown", e => {
      if (e.key === "Enter") login();
    })
  );

  ["newPassword","newPassword2"].forEach(id =>
    $(id).addEventListener("keydown", e => {
      if (e.key === "Enter") changePassword();
    })
  );
}

async function api(payload) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 30000);

  try {
    const res = await fetch(API_URL, {
      method: "POST",
      headers: {"Content-Type": "text/plain;charset=utf-8"},
      body: JSON.stringify(payload),
      signal: controller.signal
    });

    let data;
    try {
      data = await res.json();
    } catch (_) {
      throw new Error("SERVER_INVALID_RESPONSE");
    }

    if (!res.ok) throw new Error(data.error || "HTTP_ERROR");
    return data;
  } catch (err) {
    if (err.name === "AbortError") {
      throw new Error("API_TIMEOUT");
    }
    throw err;
  } finally {
    clearTimeout(timer);
  }
}

async function login() {
  clearMessage("loginMsg");

  const username = $("username").value.trim();
  const password = $("password").value;

  if (!username || !password) {
    showMessage("loginMsg","Vui lòng nhập Căn cước và mật khẩu.","error");
    return;
  }

  setBusy("loginBtn",true,"Đang đăng nhập...");

  try {
    const data = await api({
      action:"login",
      username,
      password
    });

    if (!data.ok) {
      showMessage("loginMsg",friendlyError(data.error),"error");
      return;
    }

    state.token = data.token;
    state.user = data.user;

    renderUser();

    $("logoutBtn").classList.remove("hidden");

    if (data.user.mustChangePassword) {
      showOnly("changeView");
      $("newPassword").value = "";
      $("newPassword2").value = "";
      $("newPassword").focus();
    } else {
      showOnly("payView");
      await loadPayslip();
    }
  } catch (err) {
    showMessage("loginMsg",friendlyError(err.message),"error");
  } finally {
    setBusy("loginBtn",false,"Đăng nhập");
  }
}

async function changePassword() {
  clearMessage("changeMsg");

  const p1 = $("newPassword").value;
  const p2 = $("newPassword2").value;

  if (p1.length < 8) {
    showMessage("changeMsg","Mật khẩu mới phải có ít nhất 8 ký tự.","error");
    return;
  }

  if (p1 !== p2) {
    showMessage("changeMsg","Hai mật khẩu không giống nhau.","error");
    return;
  }

  setBusy("changeBtn",true,"Đang đổi mật khẩu...");

  try {
    const data = await api({
      action:"changePassword",
      token:state.token,
      newPassword:p1,
      newPassword2:p2
    });

    if (!data.ok) {
      if (data.error === "SESSION_EXPIRED") {
        sessionExpired();
        return;
      }

      showMessage("changeMsg",friendlyError(data.error),"error");
      return;
    }

    state.user = data.user;
    renderUser();

    showMessage("changeMsg","Đổi mật khẩu thành công.","success");
    showOnly("payView");
    await loadPayslip();
  } catch (err) {
    showMessage("changeMsg",friendlyError(err.message),"error");
  } finally {
    setBusy("changeBtn",false,"Đổi mật khẩu");
  }
}

async function loadPayslip() {
  if (!state.token) return;

  clearMessage("payMsg");
  $("payslip").classList.add("hidden");

  const month = $("monthSelect").value;
  if (!month) return;

  setBusy("loadBtn",true,"Đang tải...");

  try {
    const data = await api({
      action:"getPayslip",
      token:state.token,
      payMonth:month
    });

    if (!data.ok) {
      if (data.error === "SESSION_EXPIRED") {
        sessionExpired();
        return;
      }

      if (data.error === "PASSWORD_CHANGE_REQUIRED") {
        showOnly("changeView");
        return;
      }

      showMessage("payMsg",friendlyError(data.error),"error");
      return;
    }

    renderPayslip(data);
  } catch (err) {
    showMessage("payMsg",friendlyError(err.message),"error");
  } finally {
    setBusy("loadBtn",false,"Xem phiếu");
  }
}

function renderUser() {
  if (!state.user) return;

  $("fullName").textContent = state.user.fullName || "";
  $("employeeMeta").textContent =
    `Mã NV: ${state.user.employeeId || ""} • CCCD: ${state.user.username || ""}`;
}

function renderPayslip(data) {
  const e = data.employee || {};
  const p = data.payslip || {};

  const money = v => {
    const n = Number(v);
    return Number.isFinite(n)
      ? n.toLocaleString("vi-VN") + " ₫"
      : "0 ₫";
  };

  const qty = v => {
    const n = Number(v);
    return Number.isFinite(n)
      ? n.toLocaleString("vi-VN",{maximumFractionDigits:2})
      : "0";
  };

  const esc = v => String(v ?? "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;")
    .replace(/'/g,"&#039;");

  const rows = [
    ["Tổng thu nhập","TotalIncome","money"],
    ["Lương tháng","MonthlySalary","money"],
    ["Lương cơ bản","BasicSalary","money"],
    ["Ngày công","WorkingDays","qty"],
    ["Ngày lễ","HolidayDays","qty"],
    ["Nghỉ phép hưởng lương","PaidLeaveDays","qty"],
    ["Nghỉ không lương","UnpaidLeaveDays","qty"],
    ["Ngày nghỉ tối thiểu vùng","RegionalMinimumLeaveDays","qty"],
    ["Tiền tăng ca","OvertimePay","money"],
    ["Giờ tăng ca","OvertimeHours","qty"],
    ["Giờ ngày nghỉ","RestDayHours","qty"],
    ["Giờ tăng ca ngày nghỉ","OvertimeRestDayHours","qty"],
    ["Giờ tăng ca ban đêm ngày nghỉ","NightRestDayOvertimeHours","qty"],
    ["Giờ ngày lễ","HolidayHours","qty"],
    ["Giờ tăng ca ngày lễ","HolidayOvertimeHours","qty"],
    ["Giờ tăng ca ban đêm ngày lễ","NightHolidayOvertimeHours","qty"],
    ["Ngày làm ca đêm","NightShiftDays","qty"],
    ["Giờ tăng ca ban đêm","NightOvertimeHours","qty"],
    ["Khoản khác","OtherMoney","money"],
    ["Kỷ luật","Discipline","money"],
    ["Thâm niên 2 năm","Loyalty2Years","money"],
    ["Thâm niên 5 năm","Loyalty5Years","money"],
    ["Thâm niên 10 năm","Loyalty10Years","money"],
    ["Nhà ở","Housing","money"],
    ["Đi lại","Transportation","money"],
    ["Chuyên cần","AttendanceBonus","money"],
    ["Hoa hồng bán hàng","SalesCommissionBonus","money"],
    ["Trợ cấp nghỉ phép chưa sử dụng","SeveranceUnusedLeave","money"],
    ["BHXH","SocialInsurance","money"],
    ["BHYT","HealthInsurance","money"],
    ["BHTN","UnemploymentInsurance","money"],
    ["Thuế TNCN","PersonalIncomeTax","money"],
    ["Tạm ứng","Advance","money"],
    ["Khấu trừ khác","OtherDeductions","money"],
    ["Thực nhận","NetPay","money"]
  ];

  const body = rows.map(r => `
    <tr>
      <td>${esc(r[0])}</td>
      <td class="${r[1]==="NetPay"?"net-pay":""}">
        ${r[2]==="money" ? money(p[r[1]]) : qty(p[r[1]])}
      </td>
    </tr>
  `).join("");

  $("payslip").innerHTML = `
    <div class="payslip-head">
      <h2>Phiếu lương ${esc(data.payMonth)}</h2>
      <div class="employee-card">
        <div><b>Họ tên:</b> ${esc(e.fullName)}</div>
        <div><b>Mã NV:</b> ${esc(e.employeeId)}</div>
        <div><b>CCCD:</b> ${esc(e.citizenId)}</div>
        <div><b>Phòng ban:</b> ${esc(e.department)}</div>
        <div><b>Bộ phận:</b> ${esc(e.section)}</div>
        <div><b>Chức vụ:</b> ${esc(e.position)}</div>
      </div>
    </div>
    <div class="table-wrap">
      <table class="salary-table">
        <thead><tr><th>Nội dung</th><th>Giá trị</th></tr></thead>
        <tbody>${body}</tbody>
      </table>
    </div>
  `;

  $("payslip").classList.remove("hidden");
}

async function adminLogin() {
  clearMessage("adminMsg");

  const password = $("adminPassword").value;

  if (!password) {
    showMessage("adminMsg","Vui lòng nhập mật khẩu Admin.","error");
    return;
  }

  setBusy("adminLoginBtn",true,"Đang đăng nhập...");

  try {
    const data = await api({
      action:"adminLogin",
      password
    });

    if (!data.ok) {
      showMessage("adminMsg",friendlyError(data.error),"error");
      return;
    }

    state.adminToken = data.token;
    $("adminPanel").classList.remove("hidden");
    $("adminPassword").value = "";

    showMessage("adminMsg","Đăng nhập Admin thành công.","success");
  } catch (err) {
    showMessage("adminMsg",friendlyError(err.message),"error");
  } finally {
    setBusy("adminLoginBtn",false,"Đăng nhập Admin");
  }
}

async function adminReset() {
  clearMessage("adminMsg");

  if (!state.adminToken) {
    showMessage("adminMsg","Phiên Admin đã hết hạn.","error");
    return;
  }

  const username = $("resetUsername").value.trim();

  if (!username) {
    showMessage("adminMsg","Nhập số Căn cước nhân viên cần reset.","error");
    return;
  }

  if (!confirm("Bạn có chắc muốn reset mật khẩu nhân viên này?"))
    return;

  setBusy("resetBtn",true,"Đang reset...");

  try {
    const data = await api({
      action:"adminReset",
      token:state.adminToken,
      username
    });

    if (!data.ok) {
      if (data.error === "SESSION_EXPIRED") {
        state.adminToken = "";
        $("adminPanel").classList.add("hidden");
      }

      showMessage("adminMsg",friendlyError(data.error),"error");
      return;
    }

    $("resetUsername").value = "";

    showMessage(
      "adminMsg",
      `Reset thành công. Mật khẩu tạm thời là Mã NV: ${data.employeeId}`,
      "success"
    );
  } catch (err) {
    showMessage("adminMsg",friendlyError(err.message),"error");
  } finally {
    setBusy("resetBtn",false,"Reset mật khẩu");
  }
}

function logout() {
  state.token = "";
  state.user = null;
  state.adminToken = "";

  $("password").value = "";
  $("newPassword").value = "";
  $("newPassword2").value = "";
  $("payslip").classList.add("hidden");
  $("logoutBtn").classList.add("hidden");

  showOnly("loginView");
  $("username").focus();
}

function sessionExpired() {
  logout();
  showMessage(
    "loginMsg",
    "Phiên đăng nhập đã hết hạn. Vui lòng đăng nhập lại.",
    "error"
  );
}

function showOnly(id) {
  ["loginView","changeView","payView"].forEach(x => {
    $(x).classList.toggle("hidden",x !== id);
  });
}

function showMessage(id,text,type="") {
  const el=$(id);
  el.textContent=text;
  el.className="message "+type;
}

function clearMessage(id) {
  const el=$(id);
  el.textContent="";
  el.className="message";
}

function setBusy(id,busy,text) {
  const btn=$(id);
  if (!btn) return;

  if (!btn.dataset.originalText)
    btn.dataset.originalText=btn.textContent;

  btn.disabled=busy;
  btn.textContent=busy ? text : btn.dataset.originalText;
}

function friendlyError(code) {
  const map = {
    INVALID_CREDENTIALS:"Sai Căn cước hoặc mật khẩu.",
    PASSWORD_TOO_SHORT:"Mật khẩu phải có ít nhất 8 ký tự.",
    PASSWORD_MISMATCH:"Hai mật khẩu không giống nhau.",
    PASSWORD_CANNOT_BE_EMPLOYEE_ID:"Mật khẩu mới không được trùng Mã nhân viên.",
    SESSION_EXPIRED:"Phiên đăng nhập đã hết hạn.",
    PASSWORD_CHANGE_REQUIRED:"Bạn phải đổi mật khẩu trước khi xem phiếu lương.",
    INVALID_MONTH:"Kỳ lương không hợp lệ.",
    PAYSLIP_NOT_FOUND:"Không tìm thấy phiếu lương của kỳ này.",
    INVALID_ADMIN_CREDENTIALS:"Sai mật khẩu Admin.",
    USER_NOT_FOUND:"Không tìm thấy nhân viên hoặc nhân viên đã inactive.",
    EMPLOYEE_ID_MISSING:"Nhân viên chưa có Mã NV.",
    SERVER_INVALID_RESPONSE:"Máy chủ trả về dữ liệu không hợp lệ.",
    SERVER_ERROR:"Máy chủ xảy ra lỗi.",
    INVALID_API_KEY:"API key đồng bộ không hợp lệ.",
    EMPTY_BATCH:"Batch không có dữ liệu.",
    API_TIMEOUT:"Máy chủ phản hồi quá lâu. Vui lòng thử lại."
  };

  return map[code] || "Có lỗi xảy ra. Vui lòng thử lại.";
}

function buildMonthList() {
  const select=$("monthSelect");
  if (!select) return;

  select.innerHTML="";

  const d=new Date();
  d.setDate(1);
  d.setMonth(d.getMonth()-1);

  for(let i=0;i<12;i++){
    const m=String(d.getMonth()+1).padStart(2,"0");
    const y=d.getFullYear();
    const value=`${m}-${y}`;

    const option=document.createElement("option");
    option.value=value;
    option.textContent=`${m}-${y}`;

    select.appendChild(option);
    d.setMonth(d.getMonth()-1);
  }
}
