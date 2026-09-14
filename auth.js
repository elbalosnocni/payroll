const GAS_URL = 'https://script.google.com/macros/s/AKfycbzCHDkrlhr4ZBzZUXGQe4P6RImV4YEe-IicO2W6PHWc0Fcmm9ybl3ZGyCEa78KyCyf8/exec';

const Session = {
  get token(){return sessionStorage.getItem('payroll_token');},
  set token(v){v?sessionStorage.setItem('payroll_token',v):sessionStorage.removeItem('payroll_token');},
  get role(){return sessionStorage.getItem('payroll_role');},
  set role(v){v?sessionStorage.setItem('payroll_role',v):sessionStorage.removeItem('payroll_role');},
  clear(){sessionStorage.removeItem('payroll_token');sessionStorage.removeItem('payroll_role');}
};

async function callApi(action,payload){
  const body=Object.assign({action:action},payload||{});
  if(Session.token) body.token=body.token||Session.token;
  try{
    const res=await fetch(GAS_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(body),redirect:'follow',cache:'no-store'});
    const text=await res.text();
    let json;
    try{json=JSON.parse(text);}catch(_){return {ok:false,error:'Máy chủ không trả JSON. HTTP '+res.status+'. Phản hồi: '+text.slice(0,300),code:'BAD_RESPONSE'};}
    return json;
  }catch(e){return {ok:false,error:'Không thể kết nối GAS: '+e.message,code:'NETWORK_ERROR'};}
}
function showError(el,message){el.textContent=message;el.classList.add('show');}
function hideError(el){el.textContent='';el.classList.remove('show');}

document.addEventListener('DOMContentLoaded',()=>{
  const loginForm=document.getElementById('login-form'), loginError=document.getElementById('login-error'), loginBtn=document.getElementById('login-btn');
  const cpForm=document.getElementById('change-password-form'), cpError=document.getElementById('change-password-error'), overlay=document.getElementById('change-password-overlay');
  if(Session.token) enterApp(Session.role);
  loginForm.addEventListener('submit',async e=>{
    e.preventDefault();hideError(loginError);loginBtn.disabled=true;loginBtn.textContent='Đang đăng nhập...';
    const res=await callApi('login',{cccd:document.getElementById('login-cccd').value,password:document.getElementById('login-password').value});
    loginBtn.disabled=false;loginBtn.textContent='Đăng nhập';
    if(!res.ok){showError(loginError,res.error||'Đăng nhập thất bại.');return;}
    Session.token=res.data.token;Session.role=res.data.role;enterApp(res.data.role);
    if(res.data.mustChangePassword) overlay.classList.add('show');
  });
  cpForm.addEventListener('submit',async e=>{
    e.preventDefault();hideError(cpError);
    const res=await callApi('changePassword',{oldPassword:document.getElementById('cp-old').value,newPassword:document.getElementById('cp-new').value,confirmPassword:document.getElementById('cp-confirm').value});
    if(!res.ok){showError(cpError,res.error||'Đổi mật khẩu thất bại.');return;}
    overlay.classList.remove('show');cpForm.reset();
  });
  document.querySelectorAll('.logout-btn').forEach(b=>b.addEventListener('click',async()=>{Session.clear();location.reload();}));
});
function enterApp(role){
  document.getElementById('login-shell').style.display='none';document.getElementById('app-shell').classList.add('active');
  document.getElementById('admin-nav-btn').style.display=role==='admin'?'block':'none';
  document.getElementById('payroll-nav-btn').style.display=role==='admin'?'block':'none';
  if(typeof initPayrollView==='function')initPayrollView();
}
