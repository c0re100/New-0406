<%@ LANGUAGE = VBScript %>
<!--#include file="data.asp" -->
<%
news_type=request("type")
if (news_type="") then
news_type=1
end if
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<link href="../../csss.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY leftmargin="0" topmargin="0">
<CENTER>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<tr>
<td colspan="4" height="20" align="center">
<table border="0" width="100%" cellpadding="0" cellspacing="0">
<tr>
<td width="160" bgcolor="#284070" align="center" height="19"><font color="#FFFFFF">靈劍江湖管理系統</font></td>
<td bgcolor="#4880a8" height="19">
</td></tr></table></td></tr></table>
<TABLE border="0" width="778" cellpadding="0" cellspacing="0">
<TR>
<TD width="160" valign="top" bgcolor="#4880a8">
<!--#include file="Menu.asp" -->
<br>
<BR>
</TD>
<TD valign="top">
<TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
<TR>
<TD height="21" bgcolor="#b8d4e8">&nbsp;管理 → <%=arr_news_type(news_type,1)%></TD>
</TR>
<TR>
<TD>
<TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
<TR>
<TD align="right" valign="middle">&nbsp;</TD>
<TD width="40" height="24">&nbsp;</TD>
</TR>
</TABLE>
</TD>
</TR>
<TR>
<TD align="center">
<!--#include file="../inc/function.inc.asp" -->
<table width="98%" cellpadding="4" cellspacing="1" border="0" bgcolor="#FFFFFF" class="FormTable">
<%
id=request("id")

sql="select * from hc_news where id="&id
Set rs=Server.CreateObject("ADODB.Recordset")
rs.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText

if NOT rs.EOF then
%>
<tr>
<td class="FormTd1" align="center" width="100">標　 題：</td>
<td bgcolor="#f3f8fc" class="FormTd2"><%=rs("title")%></td>
</tr>
<tr>
<td class="FormTd1" align="center" width="100">所屬類別：</td>
<td bgcolor="#f3f8fc" class="FormTd2"><%=arr_news_type(rs("type"),1)%></td>
</tr>
<tr>
<td class="FormTd1" align="center" width="100">新聞內容：</td>
<td bgcolor="#f3f8fc" class="FormTd2"><%=HtmlOut(rs("content"))%></td>
</tr>
<tr>
<td class="FormTd1" align="center" width="100">作　　者：</td>
<td bgcolor="#f3f8fc" class="FormTd2"><%=rs("author")%></td>
</tr>
<tr>
<td class="FormTd1" align="center" width="100">摘　　自：</td>
<td bgcolor="#f3f8fc" class="FormTd2"><%=rs("news_form")%></td>
</tr>
<tr>
<td class="FormTd1" align="center" width="100" height="30">狀　　態：</td>
<td height="30" class="FormTd2">
<%
if rs("hot_sign")=1 then
response.write "推薦新聞"
end if
%></td>
</tr>
<tr>
<td class="FormTd1" align="center" width="100" height="30">&nbsp;</td>
<td align="center" height="30" class="FormTd2">
<input class="button" name="Submit2" type="submit" value="‧返回‧" onclick="JavaScript:window.history.back()">
</td>
</tr>
<%
else
echo "沒有數據！"
end if
rs.close
%>
</table></TD></TR></TABLE></TD></TR></TABLE></CENTER></BODY></HTML>
