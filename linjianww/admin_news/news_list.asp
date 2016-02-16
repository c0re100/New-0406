<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../../error.asp?id=900"%>
<!--#include file="data.asp" -->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
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
<TABLE width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
<TR>
<TD colspan="4" height="20" align="center">
<TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
<TR>
<TD width="160" bgcolor="#284070" align="center" height="20"><FONT color="#FFFFFF">靈劍江湖管理系統</FONT></TD>
<TD bgcolor="#4880a8" height="20">
</TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>
<TABLE border="0" width="778" cellpadding="0" cellspacing="0">
<TR>
<TD width="160" valign="top" bgcolor="#4880a8">
<!--#include file="Menu.asp" -->
</TD>
<TD valign="top">
<TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
<TR>
<TD height="21" bgcolor="#b8d4e8">&nbsp;信息管理 → <%=arr_news_type(news_type,1)%></TD>
</TR>
<TR>
<TD align="center">
<TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
<TR>
                  <TD align="right"><IMG height="6" src="../../jh_pic/01.gif" width="7" hspace="6" align="bottom"></TD>
<TD width="65"><A href="news_add.asp?type=<%=news_type%>">增加公告</A></TD>
<TD align="right" width="20">&nbsp;</TD>
<TD width="40" height="24">&nbsp;</TD>
</TR>
</TABLE>

              <table border=0 cellpadding=2 cellspacing=0 width="98%" align="center" class="ListTable">
                <!--#include file="../inc/function.inc.asp" -->
<%
taskid=request("taskid")
action=request("action")
if (action="search") then
end if
Set rs=Server.CreateObject("ADODB.Recordset")
rs.PageSize =pagesize
sql="select * from hc_news where type='"&news_type&"' order by filltime desc"
rs.Open sql,Conn,1,1

maxpage=rs.PageCount
result_num=rs.RecordCount

if result_num<=0 then
if search="" then
word="目前還沒有記錄!"
else
word="沒有查到符合條件的記錄!"
end if
else
maxpage=rs.PageCount
page=request("page")

if Not IsNumeric(page) or page="" then
page=1
else
page=cint(page)
end if
if page<1 then
page=1
elseif  page>maxpage then
page=maxpage
end if
rs.AbsolutePage=Page
end if
%>
                <TR align="center" bgcolor="#CCCCCC"> 
                  <TD width="12" class="TableTitle">&nbsp;</TD>
                  <TD class="TableTitle">標題</TD>
                  <TD width="150" class="TableTitle">發佈日期</TD>
                  <TD width="100" class="TableTitle">操作</TD>
</TR>
<%
if (result_num>0) then
for i=1 to pagesize
j=i mod 2
%>
                <TR bgcolor="#EAEAEA"> 
                  <TD width="12" height="20"  Class="ListTd<%=j%>"><img src="../../jh_pic/01.gif" width="7" height="6"></TD>                  <TD height="20"  Class="ListTd<%=j%>"><a href="news_display.asp?id=<%=rs("id")%>"><%=rs("title")%></a></TD>
                  <TD width="150" height="20" align="center"  Class="ListTd<%=j%>"><%=FormatDT(rs("filltime"),8)%></TD>
                  <TD width="100" height="20" align=center  Class="ListTd<%=j%>"> 
                    <A href="news_modify.asp?id=<%=rs("id")%>">修改</A> <A href="news_act.asp?action=del_row&id=<%=rs("id")%>" onclick="return confirm('是否確認刪除  ？');">刪除</A> 
                  </TD>
</TR>
<%
rs.movenext
if rs.EOF then Exit For
next
end if

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<TABLE border="0" width="98%" cellpadding="0" cellspacing="0" dwcopytype="CopyTableCell">
<tr>
<td class="ListTableBottom" align="right">
<%
prop="s=&type="&news_type

maxpage=int(maxpage)
page=int(page)
call ShowLastNext (maxpage,page,prop)
%>
</td>
</tr>
</TABLE>
</TD>
</TR>
</TABLE>
</TD>
</TR>
</TABLE>
</CENTER>
</BODY>
</HTML>
