<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=request("a")
a0=trim(request.form("a"))
a1=trim(request.form("a1"))
b=trim(request.form("b"))
c=trim(request.form("c"))
d=trim(request.form("d"))
e=trim(request.form("e"))
f=trim(request.form("f"))
id=request("id")
num=trim(Request.Form("select"))
select case num
case "aa"
num="*2"
case "bb"
num="*4"
case "cc"
num="*6"
case "dd"
num="*8"
case "ee"
num="/2"
case "ff"
num="/4"
case "gg"
num="/6"
case "hh"
num="/8"
case else
%>
<script language="vbscript">
alert("�W�X�d��!")
history.back(1)
</script>
<%
end select

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if a="m" then
rs.open "SELECT id FROM �Ǫ��� where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�藍�_�A�S���өǪ��I"
	response.end
end if

	nameid=int(abs(request("id")))
	conn.Execute "update �Ǫ��� set  �Ǫ�='" & a0 & "',����=����" & num & ", ���s=���s" & num & ",��O=��O" & num & " where id="&nameid
end if
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�Ǫ��ק�')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
