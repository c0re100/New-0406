<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
 
<%if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
nickname=info(0)
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
who=Trim(Request.QueryString("who"))
if who="" then who=nickname
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
chatroombgimage=ljjh_chatroombgimage
chatroombgcolor=ljjh_chatroombgcolor
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT grade,���� FROM �Τ� where �m�W='" & info(0) & "'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���O����H���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
else
pai=rs("����")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>�ư����x</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="ccss.css" rel="stylesheet" type="text/css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed"  background=bk.jpg>
<div align="center">�ư����x<br>
�԰����Ťj��50��<br>
    ��O�Ȩ�20000000�� <br>
    �����ﹳ���Ťp��20��<br>
<form method=POST action='juezhao2ok.asp'>
�����ﹳ
<select name="to1">
<%
for i=0 to x
if show(i)<>info(0) then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
next
%>
</select>
    <br>
    <br>
     <input type=submit value="�I�i����" name=submit>
</form>
<form method=POST action='juezhao2ok.asp'>
    <select name="to1" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<<%
if session("nowinroom") = 1 then
for gw=5 to 20
gw=gw+3
gw1=gw&"�ũǪ�"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=20 to 150
gw=gw+15
gw1=gw&"�ũǪ�"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=180 to 390
gw=gw+30
gw1=gw&"�ũǪ�"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=490 to 780
gw=gw+50
gw1=gw&"�ũǪ�"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
<%
if session("nowinroom") = 1 then
for gw=801 to 1100
gw=gw+49
gw1=gw&"�ũǪ�"
%>
<option value="<%=gw1%>"><%=gw1%></option>
<%next
end if%>
    </select>
    <br>
    <input type=submit value="�I�i����" name=submit2>
  </form>
  
</div>
</body>
</html>
<%end if%>