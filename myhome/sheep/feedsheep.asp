<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
info=Session("info")
%>
<!--#include file="data2.asp"-->
<html>
<head>
<title>培養寵物</title>
<style type="text/css"></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" text="#000000" leftmargin="0" topmargin="0">
<table border="1" width="303" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000">
<tr>
<td width="371" height="21">&nbsp;</td>
</tr>
<tr>
<td width="371" height="137">
<%
rs.open"select 名字,說明,寵物類型 from 寵物狀態 where user='"&info(0)&"'",conn,1,1
if rs.bof then
rs.close
response.write "您還沒有購物呢！快去寵物商店買一隻吧"%>
<a href="indexsheep.asp">寵物商店</a>
<%else
%>
<form method="POST" action="checksheep.asp?nick=<%=rs("名字")%>">
<table border="0" width="227" cellspacing="1"
cellpadding="0" height="89" align="center">
<tr>
            <td class="p2" width="100%" height="60"> 
              <div align="center"><font color="#0000FF">□-你寵物的名字：<%=rs("名字")%> 
                <img src="image/<%=rs("說明")%>.gif" border="0" alt="<%=rs("寵物類型")%> " align="absmiddle"></font><font color="#FFCC00"> 
                </font> </div>
</td>
</tr>
<tr>
<td class="p3" width="100%" height="27">
<p align="center">
<input type="submit"
value="進入培養模式" name="B1"
style="font-family: 新細明體; font-size: 12px">
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<%
rs.close
end if
conn.close
%>
</body>

</html>