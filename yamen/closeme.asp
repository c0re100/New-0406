<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("ljjh_jm")=session("ljjh_jm")+1
if session("ljjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
inroom=trim(Request.form("chat"))
if regjm<>regjm1 then
%>
<script language=vbscript>
alert ("�A��J���{�ҽX�����T�A���ӿ�J:<%=regjm%>")
location.href = "javascript:history.back()"
</script>
<%
response.end
end if
name=request.form("name")
pass=trim(request.form("pass"))
if name="" or pass="" then
%>
<script language=vbscript>
alert "�O���O�Q�}���F�H�s�j�W�M�f�O�������H"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("ljjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then exit for
next 
if ydl=0 then
%>
<script language=vbscript>
alert "�A�b�d����r�H�A�èS���d�b����̡I�ݬݬO���O��ܿ��F�I"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
inroom=i
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select �K�X from �Τ� where �m�W='" & name & "' and �K�X='" & pass & "' "
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language=vbscript>
alert "���S���d���H�����o�ӤH�ڡH"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
if trim(pass)<>rs("�K�X") then%>
<script language=vbscript>
alert "�K�X����ڡA�|���|�O���F�H"
location.href = "javascript:history.back()"
</script>
<%
Response.End
end if
%>
<script language=vbscript>
alert "OK,�A�w�g�q�L�F�ڭ̪����ҡI"
</script>
<%
nickname=name
userip=Request.ServerVariables("REMOTE_ADDR")
kickname=name
kickwhy="�ڤ��p�u���u�F�A�S��k�I"
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i���u�޲z�j</font>"&nickname&"���������F�ۤv���p���Ѥ@�}(�ڱ��u�F�A�S��k)!" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Application.Lock
onlinelist=Application("ljjh_onlinelist"&inroom)
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(kickname) then
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
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&inroom)=listnull
else
Application("ljjh_onlinelist"&inroom)=newonlinelist
end if
Application("ljjh_useronlinename"&inroom)=useronlinename
Application("ljjh_chatrs"&inroom)=onliners
else
Application.UnLock
end if
%>
<html>
<head>
<title>���u�޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../css.css" rel=stylesheet>
</head>
<body class=p150 background="../images/8.jpg">
<div align="center">
  <table width="360" border="1" cellpadding="10" cellspacing="13" background="../images/12.jpg">
    <tr bgcolor="#FFFFFF">
      <td background="../images/7.jpg"> 
        <div align="center"><font color="#FF6600">�i �� �u �� �z �j</font> </div>
<b><br>
 �ާ@���� </b> <br>
�{�b�A�n�H�I��^�i��n���F�A�U���@�w�n�`�N�ϥΡI�@��10���������ܷ|���u���A�n�`�N�I <br>
<br>
<table border="1" cellspacing="0" cellpadding="10" bordercolorlight="ff6600" bordercolordark="#FFFFFF" align="center">
<form>
<tr>
<td> <%=sj%><br>
<%=nickname&"("&userip&")"%><br>
�{�b�A�w�g�i�H�ϥΤF�I
<div align="center">
<input type="button" value="�������f" onClick="window.close()"
name="button">
</div>
</td>
</tr>
</form>
</table>

</td>
</tr>
</table>

</div>
</body>
</html> 