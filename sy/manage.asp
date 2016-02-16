<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then Response.Redirect "../error.asp?id=439"
dim conn,rs,userconn,users
username=ljjh_nickname
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
<style type=text/css><!--td {  font-family: 新細明體; font-size: 9pt}body {  font-family: 新細明體; font-size: 9pt}select {  font-family: 新細明體; font-size: 9pt}A {text-decoration: none; font-family: "新細明體"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "新細明體"; font-size: 9pt} .big {  font-family: 新細明體; font-size: 12pt}
--></style>

</head>
<body leftmargin="0" topmargin="0" bgcolor="#3a4b91" background="../linjianww/0064.jpg" text="#000000">
<table width="650" border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#0066CC">
<tr>
<td colspan="2" height="26">
      <div align="center"><font size="+2" color="#FFFFFF">申冤狀紙一覽</font><font size="+2">&nbsp;</font></div>
</td>
</tr>
</table>

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
  <table width="650" border="1" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#99CCFF">
    <tr>
      <td width="261"><font size="2" color="#000000">○ 共有<%=Totalcount%>張狀紙，共<%=TotalPages%>頁，目前為第<%=CurPage%>頁</font></td>
      <td align="right" width="375"><font size="2" color="#000000"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=manage.asp?owner=<%=owner%>>[刷新]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[上一頁]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[下一頁]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=1>[首頁]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[尾頁]</a>&nbsp;</font></td>
</tr>
</table>
  <table width=650 border=1 cellspacing=0 cellpadding=1 align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
    <tr height=22 bgcolor="#FF6600"> 
      <td height="25" width="264"><font size="2"><b><font color="#000000">&nbsp;&nbsp;&nbsp;</font></b><font color="#FFFFFF">狀紙標題</font></font></td>
      <td align="center" colspan="2" height="25"><font size="2" color="#FFFFFF">受害者</font></td>
      <td align="center" height="25" width="91"><font size="2" color="#FFFFFF">被告者 
        </font></td>
      <td align="center" height="25" width="179"><font size="2" color="#FFFFFF">審判結果</font></td>
</tr>
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (II<RS.PageSize)
p=p+1%>
    <tr> 
      <td height="25" width="264"><font size="2"><a href=rootdisp.asp?ID=<%=rs("id")%>><%=rs("標題")%></a> 
        </font></td>
<td colspan="2" height="25"><b><font size="2"><%=left(rs("原告"),10)%></font></b></td>
      <td width="91" height="25"><b><font size="2"><%=left(rs("被告"),10)%></font></b></td>
      <td width="179" height="25"><b><font size="2"> 
        <%if rs("結果")="N" then%>
        未審理 
        <% end if %>
        <%if  rs("結果")="0" then %>
        不接受 
        <% end if %>
        <% if  rs("結果")="1" then %>
        接受 
        <% end if %>
        <a href="del.asp?ID=<%=rs("id")%>">刪除</a> </font></b></td>
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
  <table width="650" border="1" cellspacing="0" cellpadding="3" align="center" bgcolor="#99CCFF" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
<td>
<div align="center">○ 頁次： <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%>
，每頁：<font color="#CC0000"><%=RS.PageSize%></font>帖 </div>
</td>
<td align="right">
<div align="center">頁數： <a href="manage.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><<</a>
<% For I=StartPageNum to EndPageNum
if I<>CurPage then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a>
<% else %> <font color="#CC0000"><b><%=I%></b></font> <% end if %> <% Next %>
<% if EndPageNum<RS.Pagecount then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[更多...]</a>
<%end if%> </div>
</tr>
</table>
<%else%> <br>
  <div align="center"><font color="#000000">暫時無狀紙，請</font><font color="#FF0001"><a href="index.asp">添加新狀紙</a></font> 
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