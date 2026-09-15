/**
 * auth.js
 * -----------------------------------------------------------------------
 * - Địa chỉ GAS_URL: dán URL Web App đã deploy vào đây.
 * - Toàn bộ gọi API đi qua callApi() bên dưới.
 * - Token phiên đăng nhập lưu trong sessionStorage (mất khi đóng tab) để
 *   giảm rủi ro so với localStorage, nhưng vẫn tồn tại khi refresh trang.
 *
 * LƯU Ý: file này chạy trên trình duyệt thật khi deploy lên GitHub Pages
 * (không phải trong khung xem trước của Claude) - sessionStorage hoạt động
 * bình thường ở môi trường triển khai thực tế.
 * -----------------------------------------------------------------------
 */

const GAS_URL = 'https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec';

const Session = {
  get token() { return sessionStorage.getItem('payroll_token'); },
  set token(v) { v ? sessionStorage.setItem('payroll_token', v) : sessionStorage.removeItem('payroll_token'); },
  get role() { return sessionStorage.getItem('payroll_role'); },
  set role(v) { v ? sessionStorage.setItem('payroll_role', v) : sessionStorage.removeItem('payroll_role'); },
  clear() { sessionStorage.removeItem('payroll_token'); sessionStorage.removeItem('payroll_role'); }
};

/**
 * Gọi API GAS. Dùng Content-Type: text/plain để tránh trình duyệt gửi
 * preflight OPTIONS (Apps Script Web App không xử lý OPTIONS), body vẫn là
 * JSON hợp lệ và được doPost() ở backend JSON.parse bình thường.
 */
async function callApi(action, payload) {
  const body = Object.assign({ action: action }, payload || {});
  if (Session.token) body.token = body.token || Session.token;

  let res;
  try {
    res = await fetch(GAS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body)
    });
  } catch (networkErr) {
    return { ok: false, error: 'Không thể kết nối tới máy chủ. Kiểm tra lại kết nối mạng.', code: 'NETWORK_ERROR' };
  }

  let json;
  try {
    json = await res.json();
  } catch (parseErr) {
    return { ok: false, error: 'Phản hồi từ máy chủ không hợp lệ.', code: 'PARSE_ERROR' };
  }
  return json;
}

function showError(el, message) {
  el.textContent = message;
  el.classList.add('show');
}
function hideError(el) {
  el.classList.remove('show');
  el.textContent = '';
}

document.addEventListener('DOMContentLoaded', () => {
  const loginForm = document.getElementById('login-form');
  const loginError = document.getElementById('login-error');
  const loginBtn = document.getElementById('login-btn');

  const changePwForm = document.getElementById('change-password-form');
  const changePwError = document.getElementById('change-password-error');
  const changePwOverlay = document.getElementById('change-password-overlay');
  const changePwSkip = document.getElementById('change-password-skip');

  // Nếu đã có token trong session, thử vào thẳng ứng dụng
  if (Session.token) {
    enterApp(Session.role);
  }

  loginForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    hideError(loginError);
    loginBtn.disabled = true;
    loginBtn.textContent = 'Đang đăng nhập...';

    const cccd = document.getElementById('login-cccd').value.trim();
    const password = document.getElementById('login-password').value;

    const res = await callApi('login', { cccd, password });

    loginBtn.disabled = false;
    loginBtn.textContent = 'Đăng nhập';

    if (!res.ok) {
      showError(loginError, res.error || 'Đăng nhập thất bại.');
      return;
    }

    Session.token = res.data.token;
    Session.role = res.data.role;

    if (res.data.mustChangePassword) {
      changePwOverlay.classList.add('show');
      // Vẫn cho vào app phía sau lớp modal (modal chặn thao tác) để sau khi
      // đổi mật khẩu xong không cần load lại trang.
      enterApp(res.data.role);
    } else {
      enterApp(res.data.role);
    }
  });

  changePwForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    hideError(changePwError);

    const oldPassword = document.getElementById('cp-old').value;
    const newPassword = document.getElementById('cp-new').value;
    const confirmPassword = document.getElementById('cp-confirm').value;

    const res = await callApi('changePassword', { oldPassword, newPassword, confirmPassword });
    if (!res.ok) {
      showError(changePwError, res.error || 'Đổi mật khẩu thất bại.');
      return;
    }

    changePwOverlay.classList.remove('show');
    changePwForm.reset();
  });

  // Cho phép admin bỏ qua đổi mật khẩu ngay lúc này (không khuyến khích,
  // nhưng đôi khi cần thiết) - vẫn phải đổi ở lần đăng nhập sau vì
  // mustChangePassword chỉ được set false khi đổi mật khẩu thành công.
  if (changePwSkip) {
    changePwSkip.addEventListener('click', () => {
      changePwOverlay.classList.remove('show');
    });
  }

  const logoutBtns = document.querySelectorAll('.logout-btn');
  logoutBtns.forEach(btn => btn.addEventListener('click', () => {
    Session.clear();
    location.reload();
  }));
});

function enterApp(role) {
  document.getElementById('login-shell').style.display = 'none';
  document.getElementById('app-shell').classList.add('active');

  const adminNavBtn = document.getElementById('admin-nav-btn');
  const payrollNavBtn = document.getElementById('payroll-nav-btn');
  if (role === 'admin') {
    adminNavBtn.style.display = 'block';
    payrollNavBtn.style.display = 'block';
  } else {
    adminNavBtn.style.display = 'none';
    payrollNavBtn.style.display = 'none';
  }

  if (typeof initPayrollView === 'function') initPayrollView();
}
