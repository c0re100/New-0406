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
{alert("�п�J�z���m�W");
s.fromname.focus();return false}
if(s.fromemail.value=="")
{alert("�п�J�z���H�c");
s.fromemail.focus();return false}
if (s.fromemail.value.length > 0 && !s.fromemail.value.match( /^.+@.+$/ ) ) {
alert("�z���H�c��J���~!�Э��s��J!");
s.fromemail.focus();return false;}
if(s.sendtoname.value=="")
{alert("�п�J�z�n�ͪ��m�W");
s.sendtoname.focus();return false}
if(s.sendtoemail.value=="")
{alert("�п�J�z�n�ͪ��H�c");
s.sendtoemail.focus();return false}
if (s.sendtoemail.value.length > 0 && !s.sendtoemail.value.match( /^.+@.+$/ ) ) {
alert("�z�n�ͪ��H�c��J���~!�Э��s��J!");
s.sendtoemail.focus();return false;}

s.submit.disabled=true
s.reset.disabled=true
s.submit.value="�еy��..."
}
//-->
<!--
function checkhosting(){
var s=document.hosting

if(s.dateline.value=="")
{alert("�п�ܯ��δ���");
s.dateline.focus();return false}
if(s.name.value=="")
{alert("�п�J�z���m�W");
s.name.focus();return false}
if(s.email.value=="")
{alert("�п�J�z���H�c");
s.email.focus();return false}
if (s.email.value.length > 0 && !s.email.value.match( /^.+@.+$/ ) ) {
alert("�z���H�c��J���~!�Э��s��J!");
s.email.focus();return false;}
if(s.qq.value=="")
{alert("�п�J�zQQ���X");
s.qq.focus();return false}
if(s.tel.value=="")
{alert("�п�J�A���pô�q��");
s.tel.focus();return false}
if(s.pay.value=="")
{alert("�п�ܥI�ڤ覡");
s.pay.focus();return false}

s.submit.disabled=true
s.reset.disabled=true
s.submit.value="�еy��..."
}
//-->
<!--
function checkbuy(){
var s=document.buy

if(s.name.value=="")
{alert("�п�J�z���m�W");
s.name.focus();return false}
if(s.email.value=="")
{alert("�п�J�z���H�c");
s.email.focus();return false}
if (s.email.value.length > 0 && !s.email.value.match( /^.+@.+$/ ) ) {
alert("�z���H�c��J���~!�Э��s��J!");
s.email.focus();return false;}
if(s.qq.value=="")
{alert("�п�J�zQQ���X");
s.qq.focus();return false}
if(s.tel.value=="")
{alert("�п�J�A���pô�q��");
s.tel.focus();return false}
if(s.pay.value=="")
{alert("�п�ܥI�ڤ覡");
s.pay.focus();return false}

s.submit.disabled=true
s.reset.disabled=true
s.submit.value="�еy��..."
}
//-->