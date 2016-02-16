<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
nickname=info(0)
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
who=Trim(Request.QueryString("who"))
if who="" then who=nickname
show=Split(Trim(useronlinename)," ",-1)
x=UBound(show)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT grade,門派 FROM 用戶 where 姓名='" & info(0) & "'"
Set Rs=conn.Execute(sql)
pai=rs("門派")
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<HTML>
<HEAD>
<META http-equiv="Content-Style-Type" content="text/css">
<TITLE>尋找水晶</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></HEAD>
<BODY bgcolor="#ffffff" background="../bk.jpg">
<TABLE border="0" cellpadding="0" cellspacing="0" width="142">
<TR> 
<TD align="center" height="10" width="20" valign="top"><IMG src="tb-a1.gif" width="20" height="20" border="0"></TD>
<TD background="tb-a6.gif" height="10" width="102"></TD>
<TD height="10" align="center" valign="top" width="57"><IMG src="tb-a7.gif" width="20" height="20" border="0"></TD>
</TR>
<TR> 
<TD width="20" background="tb-a2.gif" height="80"></TD>
<TD width="102" height="80" align="center" valign="top" background="tb-a3.gif"><table border="0" cellspacing="0" cellpadding="1" width="102" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="287" >
<tr  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
<td width="96" height="18" bgcolor="#FFFF00"><div align="center"><font color="#CC0000"> 尋找水晶 </font></div></td></tr>
<tr>
<td height="20" bgcolor="#FF0000"> 
<div align="left"><font size="2" color="#FFFFFF">無<br>
<br>
對像：無<br>
<br>
法力：100點 <br>
<br>
效果：可能隨機變出水晶</font></div></td>
</tr>
<tr> 
<td height="4" bgcolor="#FF6633"><div align="center"> 
<form method=POST action=fs0012ok.asp>
<select name=to1 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:#000000 solid">
<%
for i=0 to x
if show(i)<>info(0) then
if show(i)<>peiou Then
%>
<option value="<%=show(i)%>"<%if CStr(show(i))=CStr(who) then Response.Write " selected"%>><%=show(i)%></option>
<%
end if
end if
next
%>
</select>
<br>
</td></tr>
<tr><td height="4" bgcolor="#FF0000"> 
<div align="center"> <br>
<input type=submit value="施展法術" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:#000000 solid">
</div></td>
</tr>
</table></TD>
<TD background="tb-a4.gif" width="57" height="80"></TD>
</TR>
<TR> 
<TD width="20" height="10"><IMG src="tb-a8.gif" width="20" height="20" border="0"></TD>
<TD background="tb-a5.gif" width="102" height="10"></TD>
<TD width="57" height="10"><IMG src="tb-a10.gif" width="20" height="20" border="0"></TD>
</TR></form></TABLE></BODY></HTML>