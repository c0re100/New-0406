<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%><%
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
nickname=info(0)
who=Trim(Request.QueryString("who"))
if who="" then who=nickname
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
rs.open "SELECT * FROM �Τ� where id="&info(9),conn

if rs.bof or rs.eof then
	response.write "�A���O���򤤤H"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("grade")>5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�̤ҩd���@�ӬO�޲z���A�T��ϥΡI');history.go(-1);</script>"
	response.end
end if
peiou=rs("�G�B")
rs.close
rs.open "SELECT * FROM �Τ� where �m�W='" & peiou & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�S���G�B�r�I�I');history.go(-1);</script>"
	response.end
end if
if rs("grade")>5 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�A�̤ҩd���@�ӬO�޲z���A�T��ϥΡI');history.go(-1);</script>"
	response.end
end if
rs.close
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>�X���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../setup.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed">
<div align="center">
<p><br>
<font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> </font></span>�X���<font size="3"><span style='font-size:9pt'> </span></font></font><font size="3"><br>
</font> <font color="#FFFFFF" size="2"><br>
�����O=�k����ˤO+�k����ˤO</font></p>
<form method=POST action='compact1.asp'>
<font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">�����ﹳ</font><br>
<select name="to1">
<%
for i=0 to x
if show(i)<>peiou and  show(i)<>info(0) Then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
next
%>
</select>
<font color="#FFFFFF" size="2"><br>
<br>
<font color="#FF0033">�X��ޤ@��</font><br>
<%
rs.open "SELECT * FROM �X��� where �m�W�k='" & info(0) & "' or �m�W�k='" & info(0) & "'"
if rs.eof and rs.bof then
	response.write "<p align='center'>�A�S������X��ޡI<br>�Цۦ��<a href=../hunyin/Stunt.asp target='_blank'>�X��S�޳B</a>�ЫءC</p>"
else%>
	<select class="smallSel" name="at" size="1"><%
	do while not rs.eof
	sel="selected"
	response.write "<option " & sel & " value='"+CStr(rs("�X���"))+"' >"+rs("�X���")+""+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	%>
	</select>
	<br>
	<br>
	</font><font color=#ffffff size=2>
	<input type=submit value="�o ��" name=submit>
	</font>
<%end if%>
<br></form>	</div></body></html>
 