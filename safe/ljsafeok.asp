<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"
Response.Buffer=true
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER")
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
name=Request.form("name")
name=trim(name)
if info(0)<>name then
Response.Write "<script language=JavaScript>{alert('�藍�_�A�W�r���šI�I');parent.history.go(-1);}</script>"
Response.End 
end if
oldpass=trim(Request.form("oldpass"))
getpass=trim(request.form("getpass"))
repass=trim(request.form("repass"))
mail=trim(Request.Form ("mail"))
repass=trim(repass)
if instr(oldpass,"'")<>0 or instr(oldpass,"|")<>0 or instr(oldpass," ")<>0 or instr(oldpass,",")<>0 then
Response.Write "<script language=JavaScript>{alert('�藍�_�A�W�r�K�X��J�����I�I');parent.history.go(-1);}</script>"
response.end
end if
if instr(getpass,"'")<>0 or instr(getpass,"|")<>0 or instr(getpass," ")<>0 or instr(getpass,",")<>0 then
Response.Write "<script language=JavaScript>{alert('�藍�_�A�K�X�O�@��J���~�I�I');parent.history.go(-1);}</script>"
response.end
end if
if instr(repass,"'")<>0 or instr(repass,"|")<>0 or instr(repass," ")<>0 or instr(repass,",")<>0 then
Response.Write "<script language=JavaScript>{alert('�藍�_�A�K�X�O�@������~�I�I');parent.history.go(-1);}</script>"
response.end
end if
if instr(mail,"'")<>0 or instr(mail,"|")<>0 or instr(mail," ")<>0 or instr(mail,",")<>0 then
Response.Write "<script language=JavaScript>{alert('�藍�_�A�l�c�榡����I�I');parent.history.go(-1);}</script>"
response.end
end if
if not(instr(mail,".")<>0 and instr(mail,"@")<>0 and len(mail)>6) then 
Response.Write "<script language=JavaScript>{alert('�藍�_�A�l�c�榡����I�I');parent.history.go(-1);}</script>"
Response.End 
end if
if oldpass="" or getpass="" or repass="" or mail="" then
message="�藍�_�I�{�ǥX��<br>�K�X�A�l�c���ର��"
elseif getpass<>repass then
message="�藍�_�I�{�ǥX��<br>�K�X�P�T�{�K�X�����@�P"
else
temppass=StrReverse(left(oldpass&"godxtll,./",10))
templen=len(oldpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
oldpass=replace(mmpassword,"'","B")
temppass=StrReverse(left(getpass&"godxtll,./",10))
templen=len(getpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
getpass=replace(mmpassword,"'","B")
'����Τ�
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.Open ("select �K�� from  �Τ� where id="&info(9)),conn
if rs("�K��")<>"" and rs("�K��")<>"1" then 
rs.Close
set rs=nothing
conn.Close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z�v�g�ӽФF�K�X�O�@�A���i�A���ӽСI�I');parent.history.go(-1);}</script>"
Response.End 
end if
rs.Close
rs.Open ("SELECT * FROM �Τ� WHERE �m�W='" & name & "' and �K�X='" & oldpass& "'"),conn
If Rs.Bof OR Rs.Eof Then
message="�藍�_�A�A���K�X����"
else
if  rs("���A")="�v" then
message="�藍�_�I�{�ǥX��<br>�A�{�b�𮧤��A����ק�K�X�I"
else
conn.Execute ("update �Τ� set �K��='" & getpass & "',�H�c='"&mail&"' where �m�W='" & name & "'")
message="���߱z���\�a�ӽФF�K�X�O�@�I"
end if
end if
ip = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
'conn.Execute ("insert into v(v1,v2,v3,v4,v5)values('"&username&"','"&sj&"','"&ip&"','"&Request.Form ("oldpass")&"','"&Request.Form ("pass")&"')")
rs.Close
set rs=nothing
conn.Close
set conn=nothing
end if
%> 
<link rel="stylesheet" href="setup.css">
<body bgcolor=#990000>
<p align="center"><font color="yellow">�� �� �� �K �X ��</font></p>
<table border=1 bgcolor="#009999" align=center width=254 cellpadding="10" cellspacing="13" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td height="180"> 
<table width=254 height="100%">
<tr align="center"> 
<td> <%=message%> 
<p align=center><a href=ljsafe.asp>�� �^</a> </p>
</td></tr>
</table>
</td>
</tr>
</table>
</body>