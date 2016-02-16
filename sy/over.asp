<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
%>
<html>
<head>
<title>申冤狀紙一覽</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../css.css">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="569" valign="bottom"><img src="../images/sy5.gif" width="58" height="23"></td>
<td width="209"><img src="../images/sy1.jpg" width="227" height="83"></td>
</tr>
<tr>
<td background="../images/sy4.gif" width="569" align="right"><img src="../images/sy3.gif" width="94" height="42"></td>
<td background="../images/sy4.gif" width="209"><img src="../images/sy2.jpg" width="227" height="42"></td>
</tr>
</table>
<br>
&nbsp;&nbsp;&nbsp;申冤狀紙一覽 <a href="manage.asp">申冤管理</a> 
<form action=over.asp method=get>
<%
set rs=server.createobject("adodb.recordset")
rs.open "SELECT * FROM 申冤  ORDER BY id DESC",Conn,1,1
if not rs.eof and not rs.bof then
RS.PageSize=15
Dim TotalPages
TotalPages = RS.PageCount

If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if

RS.AbsolutePage=CurPage

rs.CacheSize = RS.PageSize

Dim Totalcount
Totalcount =INT(RS.recordcount)
%>
<table width="650" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000">
<tr>
<td>○ 共有<font color=red><%=Totalcount%></font>張狀紙，共<%=TotalPages%>頁，目前為第<%=CurPage%>頁</td>
<td align="right"> <a href="addnew.asp"><font color="#FF0000">添加狀紙</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=over.asp?owner=<%=owner%>>[刷新]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[上一頁]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[下一頁]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=1>[首頁]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[尾頁]</a>&nbsp;</td>
</tr>
</table>
<br>
<table width=650 border=1 cellspacing=0 cellpadding=0 align="center" bordercolor="#000000" >
<tr height=22>
<td height="25" width="334" align="center"><font color="#FF6600"><b>&nbsp;&nbsp;&nbsp;</b>狀
紙 標 題 </font></td>
<td align="center" colspan="2" height="25"><font color="#FF6600">受 害 者</font></td>
<td align="center" height="25" width="100"><font color="#FF6600">被 告 者 </font></td>
<td align="center" height="25" width="93"><font color="#FF6600">審 判 結 果</font></td>
</tr>
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (II<RS.PageSize)
p=p+1%>
<tr>
<td height="25" width="334"><a href=disp.asp?ID=<%=rs("id")%>><%=rs("標題")%></a>
</td>
<td colspan="2" height="25">
<div align="center"><b><%=left(rs("原告"),10)%></b></div>
</td>
<td width="100" height="25">
<div align="center"><b><%=left(rs("被告"),10)%></b></div>
</td>
<td width="93" height="25"><b><%if rs("結果")="N" then%>未審理<% end if %><%if  rs("結果")="0" then %>不接受<% end if %><% if  rs("結果")="1" then %>接受<% end if %></b></td>
</tr>
<%rs.movenext
II=II+1
loop%>
</table>
<%StartPageNum=1
do while StartPageNum+15<=CurPage
StartPageNum=StartPageNum+15
Loop

EndPageNum=StartPageNum+14

If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount
%>
<table width="650" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td>
<div align="center">○ 頁次： <%=CurPage%>/<%=TotalPages%> ，每頁：<%=RS.PageSize%>帖
</div>
</td>
<td align="right">
<div align="center">頁數： <a href="over.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><<</a>
<% For I=StartPageNum to EndPageNum
if I<>CurPage then %> <a href="over.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a>
<% else %> <b><%=I%></b> <% end if %> <% Next %> <% if EndPageNum<RS.Pagecount then %>
<a href="over.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[更多...]</a>
<%end if%> </div>
</tr>
</table>
<%else%> <br>
<div align="center">暫時無狀紙，請<a href="addnew.asp">添加新狀紙</a>
<%
end if%>
</div>
</form>


<div align="center"> </div>
</body>
</html>
<%
rs.close
set rs=nothing
%>
<html></html> 