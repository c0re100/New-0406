<!--#include file="config.asp"-->
<!--#include file="conn.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title><%=title%></title>
<SCRIPT src="tip.js" type=text/javascript></SCRIPT>
<style type="text/css">
<!--
body,table,td,p,th {  font-size: 9pt} 
a {  text-decoration: none;color:<%=linkcolor%>} 
a:hover {  text-decoration: none} 
#tipBox {
BACKGROUND: #ffffff; BORDER-BOTTOM: black 1pt solid; BORDER-LEFT: black 1pt solid; BORDER-RIGHT: black 1pt solid; BORDER-TOP: black 1pt solid; FONT-FAMILY: 新細明體; FONT-SIZE: 9pt; POSITION: absolute; VISIBILITY: hidden; padding:2 ;Z-INDEX: 100
}
-->
</style>
</head>

<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">

<div align="center">
<table align=center cellpadding=2 cellspacing=1 width=550 border=1 bordercolor=<%=titlelightcolor%>>
<tr>
    <td align=center bgcolor=<%=titledarkcolor%>> <font color="#FFFF00"><%=title%></font> 
    </td>
</tr>
</table>
<br>
<div align="center">
<center>
<DIV id=tipBox></DIV>
<table border="0" cellpadding="0" cellspacing="0" width="550" align=center>
    <tr>
        <td valign="bottom" colspan="5"><img src="img/standtop.gif" width="500" height="30"></td>
    </tr>
    <tr>
        <td align="right" valign="top" width="50"><img src="img/standleft.gif" width="15" height="200"></td>
        <td valign="top" width="380" background="img/standbg.gif" style="padding:5">
<%
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from wish order by id desc"
rs.open sql,conn,1,1
 if not (rs.bof or rs.eof) then
    page=request.querystring("page")
    pmcount=24    '每頁顯示記錄數

    rs.pagesize=pmcount
    maxpage=rs.pagecount
    total=rs.recordcount

    if isempty(page) or cint(page)<1 or cint(page)>maxpage then
        page=1
    end if

    rs.absolutepage=page

rc=rs.pagesize
re=1
for i=1 to 4

  for x=1 to 6
   if rs.eof and rc>0 then exit for
%>
<a href ="wishshow.asp?id=<%=rs("id")%>" onmouseover="this._tip = '<font color=#808080>有<%=rs("counter")%>人看過<br><%=rs("name")%>(<%=rs("sex")%>)的願望</font><br><font color=#5a5a5a>請按此看看</font>'">
 <img src="img/<%=rs("wishtype")%>/card.gif" width=50 height=30 border=0></a>
<%
   rs.movenext
  next
 if  rs.eof and rc>0 then exit for
response.write "<br>"
next
end if
%>
</td>
        <td valign="top" width="15"><img src="img/standleft.gif" width="15" height="200"></td>
        <td valign="top" width="50"><a href="lucky.asp" target="_blank" onmouseover="this._tip = '<font color=<%=linkcolor%>>按此看看今日運程</font>'" onclick="window.open('lucky.asp','Luck','width=158,height=297,scrollbars=no'); return false;"><img src="img/luck.gif" border="0" width="50" height="100"></a><br><a href="write.asp"><img src="img/chim.gif" border="0" width="50" height="100"></a></td>
        <td valign="top" width="110"><img src="img/standleft.gif" width="15" height="200"></td>
    </tr>
</table>
</center>
</div>
<br>
<div align="center">
<table align=center cellpadding=2 cellspacing=1 width=550 border=1 bordercolor=<%=titlelightcolor%>>
    <tr>
        <td align="center" width="30%" bgcolor=<%=titledarkcolor%>>
        <p align="center"><font color="<%=titlelightcolor%>"><a href="login.asp">線上管理</font></a></p>
      </td>
        
      <td align="center" width="40%" bgcolor=<%=titledarkcolor%>><font color="<%=titlelightcolor%>"> 
        <font color="#FFFFFF">【頁數 
        <%
for num=1 to maxpage
if num=cint(page) then
response.write num 
else
response.write "<b><a href=wish.asp?page=" & num & ">" & num & "</a></b>"
end if
next
%>
        】</font></font></td>
        <td align="center" width="30%" bgcolor=<%=titledarkcolor%>>
        <p align="center"><font color="<%=titlelightcolor%>"><font color="#FFFF00">現有</font><font color="<%=linkcolor%>"><b><%=rs.recordcount%></b></font><font color="#FFFF00">個祈願</font></p>
      </td>
    </tr>
</table>
</div>
<!--#include file="copyright.asp"-->
</body>
</html>
