<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")%>
<html>
<head>
<title>�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#88AFD7">
<%
my=info(0)
if my<>"" then
Response.Write "<div align='center'><font size=-1>�v�šG"& info(7) &"</font></div>"
rs.open "select �v��,�v�ť��,�Ȩ�,�t��,���B�ɶ�,�ʧO,�p�� from �Τ� where �m�W='" & my & "'",conn
if rs("�v��")<>"" and rs("�v��")<>"�L" and rs("�v�ť��")<>"���"&Day(date()) then
	yin=int(rs("�Ȩ�")*0.05)
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"& yin &",�v�ť��='���"& Day(date()) &"' where �m�W='"& my &"'"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"& yin &" where �m�W='"& info(7) &"'"
	Response.Write "<div align='center'><font size=-1>"& my &"���ѤW��v��"& yin &"��ջȡA�ӧ��q�v�šG"& info(7) &"</font></div>"
end if
if rs("�t��")<>"�L" and rs("�ʧO")<>"�k" then
	Response.Write "<br>�A�b����w�g���B�G���B������G"& rs("���B�ɶ�") &"�A�`�N�G20�ѫ���Y�I"
end if
if rs("�t��")<>"�L" and date()-rs("���B�ɶ�")>19 and rs("�ʧO")<>"�k" and rs("�p��")="�L" then 
        Response.Write "<br>20�Ѩ�F�A<input type=button value='�n�Ĥl�F' onClick=window.open('haizi.asp?id="&info(9)&"&yhz=1')><input type=button value='���n�Ĥl' onClick=window.open('haizi.asp?id="&info(9)&"&yhz=0')>"
end if
rs.close
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
rs.open "select �U�ڤ��,�U�ڤH from �U�� where �ٶU�O��=false and �U�ڤH='"& my & "'",conn
if not(rs.BOF or rs.EOF) then
	Response.Write "<br>�A�b������˦��U�ڡA�U�ڤ���G"& rs("�U�ڤ��") &"�`�N�ٴ�,7�Ѥ��٧R��ID�I"
end if
rs.close
rs.open "select * from �U�� where �ٶU�O��=false and �U�ڤH='"& my & "' and DateDiff('d',�U�ڤ��,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	name=rs("�U�ڤH")
	conn.Execute ("update �U�� set �ٶU�O��=true where �U�ڤH='"&name&"'")
	conn.Execute ("delete * from �Τ� where �m�W='"&name&"'")
	'conn.Execute ("insert into �H�R(����,�ɶ�,����,���]) values ('" & name & "',"&sqlstr(sj)&",'���Q�U����','����ٿ��A�S���n�A�R')")
	Response.Write "�]���A7�ѨS���ٶU�ڡA����ID�Q�R���F�I"
	info(0)=""
	Session.Abandon
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
else%>
<font size=-1>�F�C�����w��z�I</font> 
<%end if
%>
<div align="center"></div>
</body>
</html>
 