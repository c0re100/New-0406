<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
grade=Int(info(2))
inthechat=Session("ljjh_inthechat")
userip=Request.ServerVariables("REMOTE_ADDR")
if info(0)="" then Response.Redirect "../error.asp?id=440"
mypai=info(5)
if info(0)="" then Response.Redirect "../error.asp?id=440"
if grade<10 and mypai<>"�x��" then Response.Redirect "../error.asp?id=482"

'���禹�ɬO�_�Q����
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT grade FROM �Τ� WHERE  id="&info(9)
rs.open sql,conn,1,3
if Not(rs.Eof and rs.Bof) then
grade=rs("grade")
if intherchat<>"1" and grade<10 then Response.Redirect "../error.asp?id=482"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
'��������
if inthechat<>"1" and  info(0)<>"�����`��" then Response.Redirect "manerr.asp?id=482"
bombname=Server.HTMLEncode(Trim(Request.Form("bombname")))
bombwhy=Server.HTMLEncode(Trim(Request.Form("bombwhy")))
logok=Trim(Request.Form("logok"))
if info(0)<>"�����`��"  then  logok="1"
if logok<>"0" then logok="1"
if bombname="" then Response.Redirect "manerr.asp?id=222"
if bombwhy="" then Response.Redirect "manerr.asp?id=224"
if len(bombwhy)>60 then bombwhy=Left(bombwhy,60)
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
cz=0
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if CStr(onlinelist(i+1))=CStr(bombname) then
cz=1
bombip=onlinelist(i+2)
end if
next
if cz=0 then Response.Redirect "manerr.asp?id=225&bombname=" & server.URLEncode(bombname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT lastkick FROM �Τ� WHERE �m�W='" & bombname & "'"
rs.open sql,conn,1,3
if Not(rs.Eof and rs.Bof) then
rs("lastkick")=sj
rs.Update
end if
rs.close
set rs=nothing
bomblog="�� <font color=red>" & bombname & "</font>(<font color=FFFF00>" & bombip & "</font>) �����I<br><font color=FFFFFF>�i��]�G" & bombwhy & "�j</font>"
if logok="1" then
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
end if
conn.close
set conn=nothing
Session("ljjh_lasttime")=sj
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
sd(194)=info(0)
sd(195)="�j�a"
sd(196)="FF0000"
sd(197)="FF0000"
sd(198)="��"
sd(199)="<font color=RED>�i�F���j</font><font color=FFFF00><font color=red>" & info(0) & "</font>�����q�q�a��<font color=red>" & bombname & "</font>���W�Ѱ�   ��]�G�i" & bombwhy & " �j</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application("ljjh_bombname")=Application("ljjh_bombname") & " " & bombname & " "
Application.UnLock
call boot(bombname)
sub boot(bombname)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(bombname) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "

if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
end sub
%><html>
<head>
<title>���u�ާ@</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF">�i���u�ާ@�j</font></h1>
</div>
<hr noshade size="1" color=009900>
<b> �ާ@���� </b> <br>
���u�w�g���X�A���W�A�N�|�ݨ��譫���a�L�F�U�h�C�A���ת��ާ@<%if logok="0" then Response.Write "<font color=red>�S��</font>"%>�Q�O���b���}�椤�C
<hr noshade size="1" color=009900>
<br>
<table border="1" cellspacing="0" cellpadding="10" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td>
<p><%=sj%></p>
<p><%=info(0)&"("&userip&")"%></p>
<p><%=bomblog%></p>
<div align="center">
<input type="button" value="��^" onclick="javascript:history.go(-2)">
</div>
</td>
</tr>
</form>
</table>
<br>
</body>
</html>
<html></html> 
