<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
who=Trim(Request.QueryString("who"))
if who="" then who=info(0)
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ���� FROM �Τ� where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���O����H���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
pai=rs("����")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>�s���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#003399" leftmargin="0" bgproperties="fixed" topmargin="0">
<div align="center">
<p><br>
<font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> </font></span>�s���<font size="3"><span style='font-size:9pt'> </span></font></font><font size="3"><br>
</font> </p>
<p class=p1 align="center"><font color="#FFFFFF" size="2">�̦h�i�s�T��<br>
�ˮ`=�T�۶ˮ`�`�M</font></p>
<form method=POST action='stunt1.asp'>
<font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">�����ﹳ</font><br>
<select name="to1">
<%
for i=0 to x
	if show(i)<>info(0) and show(i)<>peiou then
		%><option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option><%
	end if
next
%>
</select>
<font color="#FFFFFF" size="2"><br>
<br>
�s��1<br>
<select class="smallSel" name="at1" size="1">
<%
rs.close
rs.open "SELECT * FROM �Z�\ where ����='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("�Z�\"))+"' >"+rs("�Z�\")+""+chr(13)+chr(10)
rs.movenext
loop
%>
</select>
<br>
<br>
<br>
�s��2<br>
<select class="smallSel" name="at2" size="1">
<%
rs.close
rs.open "SELECT * FROM �Z�\ where ����='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("�Z�\"))+"'>"+rs("�Z�\")+""+chr(13)+chr(10)
rs.movenext
loop
%>
</select>
<br>
<br>
<br>
�s��3<br>
<select class="smallSel" name="at3" size="1">
<%
rs.close
rs.open "SELECT * FROM �Z�\ where ����='" & pai & "'",conn
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("�Z�\"))+"'>"+rs("�Z�\")+""+chr(13)+chr(10)
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
</font><br>
</form>
</div>
</body>
</html>
 