<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" or info(5)="" then Response.Redirect "../error.asp?id=440"%>
<!--#include file="dadata.asp"-->
<%
nickname=trim(info(0))
mypai=trim(info(5))
sql="select ���� from �Τ� where id="&info(9)
set rs=conn.execute(sql)
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "�A���O���򤤤H�A����i�ѥ[�b�|"
response.end
else
dj=rs("����")
if dj<2 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('�A�٬O����p���A����ѥ[�b�|');window.close();</script>"	
response.end
else

id=request("id")
sql="select �b�|�W,�֦���,���O,��O,����,����,�ƶq from �b�|�C�� where ID=" & id & " and ����='�Ҧ�' and �ѥ[��='�Ҧ�'"
Set Rs=connt.Execute(sql)
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('���u���O���A�w���A���n�èӡI�o�ˤ��n��:P');window.close();</script>"	
response.end
else
yhming=rs("�b�|�W")
yyou=rs("�֦���")
nl=rs("���O")
tl=rs("��O")
jb=rs("����")
lx=rs("����")
ls=rs("�ƶq")
if ls<1 then
sql1="delete * from �b�|�C��  where ID=" & id
connt.execute sql1
sql1="delete * from �b�|�� where �b�|�W=" & id
connt.execute sql1
connt.close
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('�A�ӱߡA�b�|�w�g����!');history.back();</script>"	
response.end
else

sql="select ID from �b�|�� where �ѥ[��='" & nickname & "'  and �b�|�W=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
sql2="insert into �b�|��(�ѥ[��,�b�|�W) values ('" & nickname & "' , " & id & " )"
connt.execute sql2
sql1="update �Τ� set ���O=���O+"&nl&",��O=��O+"&tl&",����=����+1 where �m�W='"&nickname&"' "
conn.execute sql1
sql1="update �b�|�C�� set �ƶq=�ƶq-1 where ID=" & id
connt.execute sql1
conn.close
connt.close
set rs=nothing
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
sd(199)="<font color=FFD7F4>�i�p�D�����j"&nickname&"�ѥ[�F"&yyou&"�b����j�s��"&yhming&"�U�|�檺��"&lx&"���b�|�A�W���F���֪���O�M���O�C</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<%else
connt.close
set connt=nothing
	
Response.Write "<script language=javascript>alert('�A�w�ѥ[�L�o�Ӯb�|�F�A����٨Ӱ�!');history.back();</script>"
response.end
end if%>
<%end if%>
<%end if%>
<%end if%>
<%end if%>
<html>
<title>�ѥ[�b�|���\!</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
input {border: 1px solid; font-size: 9pt; font-family: "�s�ө���"; border-color: #000000 solid}
-->
</style>

<body bgcolor="#660000">
<p>&nbsp;</p>
<table width="90%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">�ѥ[�s�b���\</font></td>
</tr>
<tr bgcolor="#66FF66">
<td align="center" height="157">
<p><font> <img src="jd/ka1.gif" width="204" height="251"></font></p>
</td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20" class="unnamed1"><b><font color="#FF3333">�A�g�L�F�@�f�T�]��|�A�ܪ��W���@�ӼҼˡA�W�[���O <%=nl%>�A��O <%=tl%>,���§���H�H�H�Ӫ���</font></b></td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20"><b><font color="#FF3333"><input  onClick="javascript:window.document.location.href='jd.asp'" value="�� �^" type=button name="back">  <INPUT onclick=window.close() type=button value=����>  </font></b></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>