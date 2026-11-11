"use strict";
/* ===== Нууц үгийн аюулгүй байдал ===== */
const pwRecovery=/type=recovery/.test(location.hash);let pwForced=false;
function openPw(forced,note){pwForced=forced;$("pwNew").value="";$("pwNew2").value="";message("pwStatus","");
 $("pwTitle").textContent=forced?"Шинэ нууц үг тохируулах":"Нууц үг солих";$("pwNote").textContent=note||"";
 $("pwCancel").classList.toggle("hidden",forced);$("pwModal").classList.remove("hidden");$("pwNew").focus()}
async function savePw(){const a=$("pwNew").value,b=$("pwNew2").value;
 if(a.length<8)return message("pwStatus","Нууц үг дор хаяж 8 тэмдэгт байна.","error");
 if(!/[A-Za-zА-Яа-яӨөҮүЁё]/.test(a)||!/\d/.test(a))return message("pwStatus","Үсэг болон тоо хоёуланг агуулна.","error");
 if(a!==b)return message("pwStatus","Давтсан нууц үг таарахгүй байна.","error");
 $("pwSave").disabled=true;message("pwStatus","Хадгалж байна...");
 try{const x=await sb.auth.updateUser({password:a,data:{pw_changed:true}});if(x.error)throw x.error;currentUser=x.data.user;
  message("pwStatus","Нууц үг амжилттай солигдлоо.","success");setTimeout(()=>$("pwModal").classList.add("hidden"),900)}
 catch(e){message("pwStatus","Алдаа: "+safeError(e),"error")}finally{$("pwSave").disabled=false}}
async function forgotPw(){const email=$("email").value.trim();if(!email)return message("loginStatus","Дээрх талбарт и-мэйлээ бичээд дахин дарна уу.","error");
 try{const x=await sb.auth.resetPasswordForEmail(email,{redirectTo:location.origin+location.pathname});if(x.error)throw x.error;
  message("loginStatus","Нууц үг сэргээх холбоосыг "+email+" хаяг руу илгээлээ. И-мэйлээ шалгана уу.","success")}
 catch(e){message("loginStatus","Алдаа: "+safeError(e),"error")}}
$("pwOpenBtn").onclick=()=>openPw(false);$("pwCancel").onclick=()=>{if(!pwForced)$("pwModal").classList.add("hidden")};
$("pwSave").onclick=savePw;$("forgotBtn").onclick=forgotPw;
{const _show2=showApp;showApp=async function(){await _show2();
  if(pwRecovery)openPw(true,"Нууц үг сэргээх холбоосоор орлоо. Шинэ нууц үгээ оруулна уу.");
  else if(!currentUser?.user_metadata?.pw_changed)openPw(true,"Аюулгүй байдлын үүднээс анхны нууц үгээ өөрийн нууц үгээр солино уу.")}}
const _lo=logout;logout=async function(){$("pwModal").classList.add("hidden");await _lo()};$("logoutBtn").onclick=()=>logout();
// Хэрэв нэвтрэлт энэ файл ачаалагдахаас өмнө дууссан бол нууц үгийн шаардлагыг шууд шалгана
if(currentUser&&!$("appSection").classList.contains("hidden")){
  if(pwRecovery)openPw(true,"Нууц үг сэргээх холбоосоор орлоо. Шинэ нууц үгээ оруулна уу.");
  else if(!currentUser.user_metadata?.pw_changed)openPw(true,"Аюулгүй байдлын үүднээс анхны нууц үгээ өөрийн нууц үгээр солино уу.")}
