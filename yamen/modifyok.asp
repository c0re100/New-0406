<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
%>
<script language=vbscript>
alert ("�A��J���{�ҽX�����T�A���ӿ�J:<%=regjm%>")
location.href = "javascript:history.back()"
</script>
<%
response.end
end if
name=Request.form("name")
name=trim(name)
'name=HTMLEncode(name)
oldpass=Request.form("oldpass")
pass=request.form("pass")
repass=request.form("repass")
repass=trim(repass)
if instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then
response.end
end if
if instr(oldpass,"'")<>0 or instr(oldpass,"|")<>0 or instr(oldpass," ")<>0 then
response.end
end if
if instr(pass,"'")<>0 or instr(pass,"|")<>0 or instr(pass," ")<>0 then
response.end
end if
if instr(repass,"'")<>0 or instr(repass,"|")<>0 or instr(repass," ")<>0 then
response.end
end if

if name="" then
message="�b�����ର��"
elseif oldpass="" or pass="" or repass="" then
message="�K�X���ର��"
elseif pass<>repass then
message="�K�X�P�T�{�K�X�����@�P"
else
temppass=StrReverse(left(oldpass&"godxtll,./",10))
templen=len(oldpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
oldpass=replace(mmpassword,"'","B")
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
'ip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")

'����Τ�
sql="SELECT ���A FROM �Τ� WHERE �m�W='" & name & "' and �K�X='" & oldpass& "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
message="�藍�_�A�A���K�X����"
else
if  rs("���A")="�v" then
message="�A�{�b�𮧤��A����ק�K�X�I"
else
sql="update �Τ� set �K�X='" & pass & "' where �m�W='" & name & "'"
conn.Execute(sql)
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& name &"','�L','�ާ@','�ﴫ�Τ�K�X�I')"
conn.close
message="���߱z���\�a�ק�F�K�X�I"
end if
end if
end if
%>
<head><LINK href="../css.css" rel=stylesheet>
</head>
<body background="../bg.gif">
<table border="1" align="center" width="360" cellpadding="10"
cellspacing="13" background="../images/12.jpg"> <tr bgcolor="#FFFFFF"> <td background="../ljimage/downbg.jpg"> 
<table width="100%"> <tr> <td> <p align="center" style="font-size:14;color:red"><%=message%></p><p align="center"><a href="#" onclick="window.close()">[ 
�� �� �� �f ]</a></p></td></tr> </table></td></tr> </table>

</body> 
<html><script language="JavaScript">                                                                  </script></html>