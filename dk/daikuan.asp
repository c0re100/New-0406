<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
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
sj=n & "-" & y & "-" & r
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select  allvalue,�Ȩ�,����,�s�� from �Τ� where id="&info(9),conn
allvalue=rs("allvalue")
bigdk=allvalue*100
yinliang=rs("�Ȩ�")
jhdenji=rs("����")
nowyl=rs("�Ȩ�")
nowck=rs("�s��")
rs.close
rs.open "select �U�ڤH from �U�� where �ٶU�O��=false and DateDiff('d',�U�ڤ��,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	do while not (rs.bof or rs.eof)
		name=rs("�U�ڤH")
		conn.Execute ("update �U�� set �ٶU�O��=true where �U�ڤH='"&name&"'")
		'conn.Execute ("delete * from �Τ� where �m�W='"& name &"'")
		conn.Execute ("insert into �H�R(����,�ɶ�,����,���]) values ('" & name & "',"&sqlstr(sj)&",'���Q�U����','����ٿ��A�S���n�A�R')")
		rs.movenext
		name=""
		info(0)=""
	loop
end if
%>
<html>
<head>
<title>�U��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="chat/READONLY/Style.css">
</head>

<body bgcolor="#FFFFFF" background="../images/8.jpg">
<p align="center"><font size="6" face="����">���򰪧Q�U��</font></p>
<p align="center">&nbsp;</p>
<form method="post" action="dk.asp">
<table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#E0E0E0">
<tr>
<td>�ӶU�H:</td>
<td><%=info(0)%></td>
</tr>
<tr>
<td>�{�b�Ȩ�G</td>
<td><%=nowyl%>��</td>
</tr>
<tr>
<td>�{�b�s��:</td>
<td><%=nowck%>��</td>
</tr>
<tr>
<td>�̤j�U�ڪ��B:</td>
<td><%=bigdk%>��</td>
</tr>
<tr>
<td>�U�ڪ��B�G</td>
<%
rs.close
rs.open "select �U�ڪ��B,�U�ڤ�� from �U�� where �U�ڤH='"&info(0)&"' and �ٶU�O��=false",conn
if rs.BOF or rs.EOF then
	%>
	<td>
	<%if jhdenji>3 then%>
		<input type="text" name="dk" size="10" maxlength="10">
	<%else%>
		<font color=red>����ާ@</font>
	<%end if%>
	</td>
	</tr>
<tr>
<td colspan="2">
<div align="center">
<%if jhdenji>3 then%>
	<input type=submit value="�U��" name="submit">
	<input type="reset" name="reset" value="�M��">
<%else%>
	<font color=red>�԰��p��3�Ŧ��򤣶U�ڵ��A�I</font>
<%end if%>
</div>
</td>
<%else%>
<td>
<%if yinliang>rs("�U�ڪ��B")*1.5 then%>
<input type="text" name="dk" size="10" maxlength="10" value='<%=rs("�U�ڪ��B")%>' readonly>
<%else%>
<font color=red>����ާ@</font>
<%end if%>
<br>�U�ڤ���G<%=rs("�U�ڤ��")%>

</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<%if yinliang>rs("�U�ڪ��B")*1.5 then%>
<input type=submit value="�ٴ�" name="submit">
<input type="reset" name="reset" value="�M��">
<%else%>
<font color=red>�A���{�������ٴ�!</font>
<%end if%>
</div>
</td>
<%end if
rs.close
set rs=nothing
conn.Close
set conn=nothing%>
</tr>
</table>
</form>
<p align="center">���������ѶU�ڡA�p�H�����U�]�԰����Ť���3�Ť���U�^<br>
�U�ڴ����O�@�ӬP��7�ѡC�U�ܻ�&quot;����ٿ��A�S���٩R&quot;�A������٪̥����N��<br>
���H�N������A<font color=red>(�N�R������ID)</font>��U��j�L�ۤ���i�I�I�I�I�I�I <br>
(���Q�U�ٿ���Ҭ��U100��A��150��A���O�ڤ߶§r�A�{�b�ȿ����r�I) </p>
</body>
</html>
<%
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
%>