<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Session("diaoyu")=true
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=FFCFCF>����</font>"& info(0) &"�b�����������G�ѩ�j�k�ӯ�,�Q�����ԤF�@��˦b�Y�W,�ٵ������k�F."
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>�Q�k�F</title></head>
<body BGCOLOR="#996699">
<div align="center"> <p>&nbsp;</p><table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="200"> 
<tr> <td bgcolor=#C6BD9B> <table> <tr> <td valign="top"> <div align="center"> 
�n�h�W,�A�ѩ�j�k�ӯ�,�Q�����ԤF�@��˦b�Y�W,�ٵ������k�F... </div></table><div align="center"><br> <input type=button value=�������f onClick='window.close()' name="button"> 
</div></td></tr> </table></div>
</body>
</html> 