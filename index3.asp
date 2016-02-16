<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
randomize timer
s=1+int(rnd*10000)%>
<%randomize timer
mysound=int(rnd*8)+1
Response.Write "<bgsound src=mid/midi" & mysound & ".mid loop=-1 volume=100>"
%>
<!--#include file="config.asp"-->
<script language="JavaScript" type="text/JavaScript">
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
</script>
<SCRIPT LANGUAGE="JavaScript">
function check(got) {
if (got.name.value=="") {
alert("請輸入用戶名 !!");
return false;
}
if (got.pass.value=="") {
alert("請輸入密碼 !!");
return false;
}
return true;
}
function openwin08(){
window.open('','ewin',"height=400,width=575,screenX=250,screenY=150,top=200,left=250,resizable=no,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");
}
function openwin07(){
window.open('','fwin',"height=360,width=570,screenX=250,screenY=150,top=100,left=250,resizable=no,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");
}
</script>
<HTML><HEAD><TITLE>魔神江湖</TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<LINK href="http://chichi.sytes.net/air/upload/images/air/001/dob.css" type=text/css rel=stylesheet>
<style>
<!--
A {
	CURSOR: url('149.cur')
}
-->
</style>

<BODY leftMargin=0 background=bg.gif topMargin=0 marginheight="0" marginwidth="0" scroll="no">
<DIV align=center>
  <TABLE width=600  border=0 align=center cellPadding=0 cellSpacing=0>
<TR>
<TD><font size="2"><img src="kk.gif" width="773" align="middle" height="186" ></font></TD>
</TR>
</TABLE>
<DIV align=center style="width: 1008; height: 580">
  <TABLE height=19 cellSpacing=0 cellPadding=0 width=771 border=0>
  <tr>
    <TD height=20 width="128">
      <DIV align=center><font size="2">
        <img src="a1.gif" width="11" height="11" border="0">
        </A></font><font size="2"><a href="jh.asp">參觀江湖</a></font></DIV></TD>
    <TD height=20 width="128">
      <DIV align=center><font size="2">
        <img src="a2.gif" width="11" height="11" border="0">
        </A></font><a href="#" onClick="window.open('yamen/joinjh.asp','rg','scrollbars=yes,resizable=yes,width=365,height=360')" >
        <font size="2">註冊帳號</font></a></DIV></TD>
    <TD width="128">
      <DIV align=center>
        <font size="2">
        <IMG height=11 
      src="a3.gif" width=11 border=0> </font> <a href="#" onClick="window.open('yamen/modify.asp','rmd','scrollbars=yes,resizable=yes,width=350,height=360')" >
        <font size="2">修改密碼</font></a></DIV></TD>
    <TD width="129">
      <DIV align=center>
        <font size="2">
        <IMG height=11 
      src="a4.gif" width=11 border=0> </font> <a href="#" onClick="window.open('yamen/disp.asp','r3g','scrollbars=yes,resizable=yes,width=360,height=310')" >
        <font size="2">帳號復活</font></a></DIV></TD>
    <TD width="129">
      <DIV align=center>
        <font size="2">
        <IMG height=11 
      src="a5.gif" width=11 border=0> </font> <a href="#" onClick="window.open('yamen/close.asp','rag','scrollbars=yes,resizable=yes,width=365,height=330')" >
        <font size="2">掉線處理</font></a></DIV></TD>
    <TD width="129">
      <DIV align=center>
        <font size="2">
        <IMG 
      height=11 src="a6.gif" width=11 border=0> </font>
        <A 
      href="javascript:window.external.AddFavorite('http://kek2006.somee.com','魔神江湖')">
        <font size="2">加入我的最愛</font></A></DIV></TD>
  </tr>
  </TABLE>
<TABLE height=150 cellSpacing=0 cellPadding=0 width=773 align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=top width=230 height=146>
      <TABLE cellSpacing=0 cellPadding=0 border=0>
        <TBODY>
        <TR>
          <TD vAlign=top align=middle width=227 height=19>
            <DIV align=center>
            <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
              <TBODY>
              <tr>
<td height="107"> 
<div align="center">
<table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#9BB1E6" bgcolor="#FFFFFF" style="border-collapse: collapse">
<form method=POST action=check.asp  onSubmit="return(check(this))">
<tr> 
<td width="100%" height="18" colspan="2">
<p align="center"><img border="0" src="a6.gif"><font color="#9BB1E6" style="font-size: 12pt" face="標楷體"><a href="index2.asp">入唔到江湖請按</a></font></td>
</tr>
<tr> 
<td width="100%" height="18" colspan="2">
　</td>
</tr>
<tr> 
<td width="100%" height="18" colspan="2">
<p align="center"> 
 
  <select name=select size="9">
");<%
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("ljjh_useronlinename"&i))," ")
	onlinenum=ubound(online)+1
	chatroomxx=split(ljjh_roominfo(i),"|")
	onlinenow=onlinenow+onlinenum
next 
%>
document.write("<option selected style='color:#000000; '>共<%=onlinenow%>人在戀愛中</option>

");<%
for i=0 to chatroomnum	
	online=split(trim(Application("ljjh_useronlinename"&i))," ")
	onlinenum=ubound(online)+1
	chatroomxx=split(ljjh_roominfo(i),"|")
%>
document.write("<option style='font-size:9pt;color:#FFFFFF;background-color:42774F'><%=chatroomxx(0)%><%=onlinenum%>人聊天</option>

");<%	if onlinenum<>0 then
useronlinename=replace(trim(Application("ljjh_useronlinename"&i))," ","</option><option style='font-size:9pt;color:#000000;'>")
%>
document.write("<option style='color:#000000; '><%=useronlinename%></option>
");<%end if
next
%>
document.write("</select></td>
</tr>
<tr> 
<td><div align="right"><font size="2" color="#9BB1E6">用戶：</font></div></td>
<td>
<input type="text" name="name" maxlength="12"  size="20" class='tinyBorder' style='border:1px solid #000000; width: 110'></td>
</tr>
<tr> 
<td height="25"> <div align="right"><font size="2" color="#9BB1E6">密碼：</font></div></td>
<td>
<input type="password" name="pass" maxlength="12"  size="20" class='tinyBorder' style='border:1px solid #000000; width: 110'></td>
</tr>
<tr> 
<td colspan="2"><div align="center">
<input type="submit" name="Submit" value="登錄江湖" class='tinyBorder' style='border:1px solid #000000; width: 60; background-color:#FFFFFF'><font size="2">
</font>
<input type="reset" name="Submit2" value="重新填寫" class='tinyBorder' style='border:1px solid #000000; width: 60; background-color:#FFFFFF'><font size="2">
</font>
</div></td>
</tr></form>
</table>
</div></td>
              </tr>
              </TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
    <TD vAlign=top align=middle width=543>
      <TABLE cellSpacing=0 width="537" border=0 cellpdding="0">
        <TBODY>
        <TR>
          <TD align=middle width="535">
            <TABLE cellSpacing=0 cellPadding=0 border=0>
              <TBODY>
              <TR>
                <TD width=118 background=bianji.gif 
                  height=23><FONT color=#ddebff size="2">&nbsp; 站 長 公 告</FONT></TD>
                <TD width=395 background=bj_b.gif height=23 
                BORDER="0"></TD>
                <TD>
                <font size="2">
                <IMG height=23 src="bj_r.gif" width=5 
                  border=0></font></TD></TR></TBODY></TABLE>
            <TABLE cellSpacing=0 cellPadding=0 border=0>
              <TBODY>
              <TR>
                <TD width=4 background=tab_l.gif></TD>
                <TD vAlign=top align=middle width=510 bgColor=#9bb1e6>
<TABLE width="100%" border="0" cellpadding="1" cellspacing="1" height="174">

<TR>
<TD bgcolor="#DDEBFF" height="171">　<iframe name="I1" width="480" height="256" src="new_news/new.asp" border="0" frameborder="0" target="_self" align="center">您的瀏覽器不支援內置框架或目前的設定為不顯示內置框架。</iframe>         
</TD>
</TR>
</TABLE>
                </TD>
                <TD width=4 background=tab_r.gif></TD></TR>
              <TR>
                <TD colSpan=3>
                <font size="2">
                <IMG height=10 
                  src="tab_foot.gif" width=518 
              border=0></font></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TR></TBODY></TABLE>
              <p align="center"><b><font size="2"></a></font></b><p align="center">
　<p align="center">　</DIV>
</DIV>
</BODY></HTML>
<!-- #include file="global.asp" -->