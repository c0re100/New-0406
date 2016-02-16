<!--
if (navigator.appName=="Netscape" && parseInt(navigator.appVersion)==4) { 
	document.onmousedown=NSclick;
	document.captureEvents(Event.MOUSEDOWN);
}
if (navigator.appName=="Netscape" && parseInt(navigator.appVersion)>=5) { 
	document.onmouseup=NSclick;
}
if (navigator.appName=="Microsoft Internet Explorer") { 
	document.oncontextmenu = new Function("return false;")
	document.onmousemove = new Function("return false;")
}
//-->
<!--
top.window.moveTo(0,0);
if (document.all) {
top.window.resizeTo(screen.availWidth,screen.availHeight);
}
else if (document.layers||document.getElementById) {
if (top.window.outerHeight<screen.availHeight||top.window.outerWidth<screen.availWidth){
top.window.outerHeight = screen.availHeight;
top.window.outerWidth = screen.availWidth;
}
}
//-->
<!--
function checksendemail(){
var s=document.send

if(s.fromname.value=="")
{alert("請輸入您的姓名");
s.fromname.focus();return false}
if(s.fromemail.value=="")
{alert("請輸入您的信箱");
s.fromemail.focus();return false}
if (s.fromemail.value.length > 0 && !s.fromemail.value.match( /^.+@.+$/ ) ) {
alert("您的信箱輸入錯誤!請重新輸入!");
s.fromemail.focus();return false;}
if(s.sendtoname.value=="")
{alert("請輸入您好友的姓名");
s.sendtoname.focus();return false}
if(s.sendtoemail.value=="")
{alert("請輸入您好友的信箱");
s.sendtoemail.focus();return false}
if (s.sendtoemail.value.length > 0 && !s.sendtoemail.value.match( /^.+@.+$/ ) ) {
alert("您好友的信箱輸入錯誤!請重新輸入!");
s.sendtoemail.focus();return false;}

s.submit.disabled=true
s.reset.disabled=true
s.submit.value="請稍候..."
}
//-->
<!--
function checkhosting(){
var s=document.hosting

if(s.dateline.value=="")
{alert("請選擇租用期限");
s.dateline.focus();return false}
if(s.name.value=="")
{alert("請輸入您的姓名");
s.name.focus();return false}
if(s.email.value=="")
{alert("請輸入您的信箱");
s.email.focus();return false}
if (s.email.value.length > 0 && !s.email.value.match( /^.+@.+$/ ) ) {
alert("您的信箱輸入錯誤!請重新輸入!");
s.email.focus();return false;}
if(s.qq.value=="")
{alert("請輸入您QQ號碼");
s.qq.focus();return false}
if(s.tel.value=="")
{alert("請輸入你的聯繫電話");
s.tel.focus();return false}
if(s.pay.value=="")
{alert("請選擇付款方式");
s.pay.focus();return false}

s.submit.disabled=true
s.reset.disabled=true
s.submit.value="請稍候..."
}
//-->
<!--
function checkbuy(){
var s=document.buy

if(s.name.value=="")
{alert("請輸入您的姓名");
s.name.focus();return false}
if(s.email.value=="")
{alert("請輸入您的信箱");
s.email.focus();return false}
if (s.email.value.length > 0 && !s.email.value.match( /^.+@.+$/ ) ) {
alert("您的信箱輸入錯誤!請重新輸入!");
s.email.focus();return false;}
if(s.qq.value=="")
{alert("請輸入您QQ號碼");
s.qq.focus();return false}
if(s.tel.value=="")
{alert("請輸入你的聯繫電話");
s.tel.focus();return false}
if(s.pay.value=="")
{alert("請選擇付款方式");
s.pay.focus();return false}

s.submit.disabled=true
s.reset.disabled=true
s.submit.value="請稍候..."
}
//-->