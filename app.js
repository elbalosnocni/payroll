const API_URL =
  'https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9yblZ3GyCEa78KyCyf8/exec';

const state = {
  token: null,
  user: null,
  adminToken: null
};

const $ = id => document.getElementById(id);

function setMessage(id, text, error=false){
  const el = $(id);
  el.textContent = text || '';
  el.style.color = error ? '#c62828' : '#237a3b';
}

async function api(body){
  const response = await fetch(API_URL,{
    method:'POST',
    headers:{
      'Content-Type':'text/plain;charset=utf-8'
    },
    body:JSON.stringify(body)
  });

  if(!response.ok){
    throw new Error('HTTP '+response.status);
  }

  return response.json();
}

function money(value){
  return Number(value || 0).toLocaleString('vi-VN') + ' ₫';
}

function qty(value){
  return Number(value || 0).toLocaleString('vi-VN',{
    minimumFractionDigits:0,
    maximumFractionDigits:2
  });
}

function salaryRow(label,value,moneyMode=true){
  return `
    <div class="salary-row">
      <span>${escapeHtml(label)}</span>
      <span class="${moneyMode?'money':''}">
        ${moneyMode ? money(value) : qty(value)}
      </span>
    </div>`;
}

function escapeHtml(value){
  return String(value ?? '')
    .replaceAll('&','&amp;')
    .replaceAll('<','&lt;')
    .replaceAll('>','&gt;')
    .replaceAll('"','&quot;')
    .replaceAll("'","&#039;");
}

/*
 * Tháng hiển thị mặc định = tháng trước.
 * Ví dụ 10/09/2026 -> 08-2026.
 */
function buildMonthList(){
  const list=[];
  const d=new Date();

  for(let i=0;i<12;i++){
    const x=new Date(d.getFullYear(),d.getMonth()-1-i,1);
    const value =
      String(x.getMonth()+1).padStart(2,'0')+'-'+x.getFullYear();
    list.push(value);
  }

  $('monthSelect').innerHTML =
    list.map(x=>`<option value="${x}">${x}</option>`).join('');
}

buildMonthList();

$('loginBtn').onclick = async () => {
  setMessage('loginMsg','Đang kiểm tra...');

  try{
    const result = await api({
      action:'login',
      username:$('username').value.trim(),
      password:$('password').value
    });

    if(!result.ok){
      setMessage('loginMsg','Số Căn cước hoặc mật khẩu không đúng.',true);
      return;
    }

    state.token=result.token;
    state.user=result.user;

    $('loginView').classList.add('hidden');
    $('logoutBtn').classList.remove('hidden');

    $('fullName').textContent=result.user.fullName;
    $('employeeMeta').textContent =
      `Mã NV: ${result.user.employeeId} | CCCD: ${result.user.username}`;

    if(result.mustChangePassword){
      $('changeView').classList.remove('hidden');
    }else{
      $('payView').classList.remove('hidden');
      await loadPayslip();
    }

  }catch(e){
    setMessage('loginMsg','Không kết nối được máy chủ.',true);
  }
};

$('changeBtn').onclick = async () => {
  const p1=$('newPassword').value;
  const p2=$('newPassword2').value;

  if(p1.length<8){
    setMessage('changeMsg','Mật khẩu mới phải có ít nhất 8 ký tự.',true);
    return;
  }

  if(p1!==p2){
    setMessage('changeMsg','Hai mật khẩu không giống nhau.',true);
    return;
  }

  try{
    const result=await api({
      action:'changePassword',
      token:state.token,
      newPassword:p1
    });

    if(!result.ok){
      setMessage('changeMsg','Không thể đổi mật khẩu.',true);
      return;
    }

    setMessage('changeMsg','Đổi mật khẩu thành công.');

    setTimeout(async()=>{
      $('changeView').classList.add('hidden');
      $('payView').classList.remove('hidden');
      await loadPayslip();
    },500);

  }catch(e){
    setMessage('changeMsg','Không kết nối được máy chủ.',true);
  }
};

$('loadBtn').onclick=loadPayslip;

async function loadPayslip(){
  setMessage('payMsg','Đang tải phiếu lương...');

  try{
    const result=await api({
      action:'getPayslip',
      token:state.token,
      month:$('monthSelect').value
    });

    if(!result.ok){
      $('payslip').classList.add('hidden');
      setMessage(
        'payMsg',
        'Không tìm thấy phiếu lương của tháng đã chọn.',
        true
      );
      return;
    }

    renderPayslip(result.payslip);
    setMessage('payMsg','');

  }catch(e){
    setMessage('payMsg','Không kết nối được máy chủ.',true);
  }
}

function renderPayslip(p){
  const e=p.employee;
  const i=p.income;
  const d=p.deductions;

  $('payslip').innerHTML=`
    <h2>Phiếu lương tháng ${escapeHtml(p.month)}</h2>
    <div class="muted">
      Xưởng: ${escapeHtml(p.factory)}
      &nbsp; | &nbsp;
      UpdatedAt: ${escapeHtml(p.updatedAt)}
    </div>

    <div class="section-title">Thông tin nhân viên</div>

    <div class="info-grid">
      <div class="info">
        <b>Họ tên</b><br>${escapeHtml(e.fullName)}
      </div>
      <div class="info">
        <b>Mã nhân viên</b><br>${escapeHtml(e.employeeId)}
      </div>
      <div class="info">
        <b>Số Căn cước</b><br>${escapeHtml(e.citizenId)}
      </div>
      <div class="info">
        <b>Phòng ban</b><br>${escapeHtml(e.department)}
      </div>
      <div class="info">
        <b>Bộ phận</b><br>${escapeHtml(e.section)}
      </div>
      <div class="info">
        <b>Chức vụ</b><br>${escapeHtml(e.position)}
      </div>
    </div>

    <div class="section-title">A. Các khoản thu nhập</div>

    ${salaryRow('Tổng các khoản thu nhập',i.totalIncome)}

    <p><b>1. Lương tháng</b></p>

    ${salaryRow('Lương tháng',i.monthlySalary)}
    ${salaryRow('Lương cơ bản',i.basicSalary)}
    ${salaryRow('Số ngày làm việc',i.workingDays,false)}
    ${salaryRow('Số ngày lễ',i.holidayDays,false)}
    ${salaryRow('Số ngày nghỉ hưởng lương',i.paidLeaveDays,false)}
    ${salaryRow('Số ngày nghỉ không hưởng lương',i.unpaidLeaveDays,false)}
    ${salaryRow('Số ngày nghỉ hưởng lương tối thiểu vùng',i.regionalMinimumLeaveDays,false)}

    <p><b>2. Lương làm ngoài giờ, ngày nghỉ, ca đêm</b></p>

    ${salaryRow('Tiền làm ngoài giờ / ngày nghỉ / ca đêm',i.overtimePay)}
    ${salaryRow('Số giờ làm ngoài giờ',i.overtimeHours,false)}
    ${salaryRow('Số giờ làm ngày nghỉ',i.restDayHours,false)}
    ${salaryRow('Số giờ làm ngoài giờ ngày nghỉ',i.overtimeRestDayHours,false)}
    ${salaryRow('Số giờ tăng ca đêm ngày nghỉ',i.nightRestDayOvertimeHours,false)}
    ${salaryRow('Số giờ làm vào ngày lễ',i.holidayHours,false)}
    ${salaryRow('Số giờ làm ngoài giờ ngày lễ',i.holidayOvertimeHours,false)}
    ${salaryRow('Số giờ tăng ca đêm ngày lễ',i.nightHolidayOvertimeHours,false)}
    ${salaryRow('Số ngày làm ca đêm',i.nightShiftDays,false)}
    ${salaryRow('Số giờ tăng ca đêm',i.nightOvertimeHours,false)}

    <p><b>3. Phụ cấp, trợ cấp, hỗ trợ, bổ sung</b></p>

    ${salaryRow('Tiền khác (công tác phí, cơm)',i.otherMoney)}
    ${salaryRow('Tiền kỷ luật',i.discipline)}
    ${salaryRow('Tiền gắn bó 2 năm',i.loyalty2Years)}
    ${salaryRow('Tiền gắn bó 5 năm',i.loyalty5Years)}
    ${salaryRow('Tiền gắn bó 10 năm',i.loyalty10Years)}
    ${salaryRow('Tiền nhà ở',i.housing)}
    ${salaryRow('Tiền đi lại',i.transportation)}
    ${salaryRow('Tiền thưởng chuyên cần',i.attendanceBonus)}
    ${salaryRow('Hoa hồng bán hàng & thưởng vượt định mức',i.salesCommissionBonus)}
    ${salaryRow('Trợ cấp thôi việc & phép năm còn lại',i.severanceUnusedLeave)}

    <div class="section-title">B. Các khoản khấu trừ</div>

    ${salaryRow('BHXH',d.socialInsurance)}
    ${salaryRow('BHYT',d.healthInsurance)}
    ${salaryRow('BHTN',d.unemploymentInsurance)}
    ${salaryRow('Thuế thu nhập',d.personalIncomeTax)}
    ${salaryRow('Tạm ứng',d.advance)}
    ${salaryRow('Các khoản khấu trừ khác',d.otherDeductions)}

    <div class="section-title">C. Lương thực lĩnh</div>

    <div class="salary-row net">
      <span>LƯƠNG THỰC LĨNH</span>
      <span>${money(p.netPay)}</span>
    </div>
  `;

  $('payslip').classList.remove('hidden');
}

$('logoutBtn').onclick=()=>{
  state.token=null;
  state.user=null;

  $('loginView').classList.remove('hidden');
  $('changeView').classList.add('hidden');
  $('payView').classList.add('hidden');
  $('logoutBtn').classList.add('hidden');
  $('payslip').classList.add('hidden');

  $('password').value='';
  $('newPassword').value='';
  $('newPassword2').value='';
};

$('adminLoginBtn').onclick=async()=>{
  setMessage('adminMsg','Đang kiểm tra...');

  try{
    const result=await api({
      action:'adminLogin',
      password:$('adminPassword').value
    });

    if(!result.ok){
      setMessage('adminMsg','Sai mật khẩu Admin.',true);
      return;
    }

    state.adminToken=result.token;
    $('adminPanel').classList.remove('hidden');
    setMessage('adminMsg','Đăng nhập Admin thành công.');

  }catch(e){
    setMessage('adminMsg','Không kết nối được máy chủ.',true);
  }
};

$('resetBtn').onclick=async()=>{
  const username=$('resetUsername').value.trim();

  if(!username){
    setMessage('adminMsg','Nhập số Căn cước cần reset.',true);
    return;
  }

  try{
    const result=await api({
      action:'adminReset',
      token:state.adminToken,
      username:username
    });

    if(!result.ok){
      setMessage(
        'adminMsg',
        'Reset thất bại: '+result.error,
        true
      );
      return;
    }

    setMessage(
      'adminMsg',
      'Reset thành công. Mật khẩu tạm thời là Mã nhân viên.'
    );

  }catch(e){
    setMessage('adminMsg','Không kết nối được máy chủ.',true);
  }
};
