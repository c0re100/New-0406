<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if Session("ljjh_inthechat")<>"1" then
Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
who=Trim(Request.QueryString("who"))
if who="" then who=info(0)
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �Τ� where �m�W='" & info(0) & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���O����H���I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if chatbgcolor="" then chatbgcolor="008888"
rs.close
%>
<html>
<head>
<title>�d������</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../setup.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#003399" leftmargin="0" bgproperties="fixed" topmargin="0">
<div align="center">
<p><br>
<font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> </font></span>�d������<font size="3"><span style='font-size:9pt'> </span></font></font><font size="3"><br>
</font> <font color="#FFFFFF" size="2"><br>
�����O=���ޯण�P�Ӥ��P</font></p>
<form method=POST action='pet1.asp'>
<font color="#FFFFFF" size="2"> </font><font size="2" color="#FF0033">�����ﹳ</font><br>
<select name="to1" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<%for i=0 to x
if show(i)<>info(0) then%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%end if
next%>
</select>
<font color="#FFFFFF" size="2"><br>
<br>
<font color="#FF0033">�d���ޯ�@��</font><br>
<%
sql="SELECT * FROM �ޯ�C�� where �֦���='" &  info(0) & "'"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
response.write "�A�S�������d���ޯ�I"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
%>
<select class="smallSel" name="at" size="1" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<%
do while not rs.eof
sel="selected"
response.write "<option " & sel & " value='"+CStr(rs("�ޯ�W��"))+"' >"+rs("�ޯ�W��")+""+chr(13)+chr(10)
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
</form>
</div>
</body>
</html>