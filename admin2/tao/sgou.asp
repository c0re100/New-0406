<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data.asp"-->
<%
sql="select * from 淘金者 where 擁有者='" & info(0) & "' and (金>0 or 銀>0 or 銅>0)"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
       set rs=nothing
       conn.close
       set conn=nothing
%>
<script language=vbscript>
  MsgBox "你還沒有礦呢!"
  window.close()
</script>
<%
response.end
end if
 jin=rs("金")*10
 yin=rs("銀")*6
 tong=rs("銅")*3
 money=jin+yin+tong
 conn.execute("update 淘金者 set 金=0,銀=0,銅=0 where 擁有者='" & info(0) & "'")
 set rs=nothing
 conn.close
 set conn=nothing
%>
<!--#include file="conn.asp"-->
<%
 conn.execute("update 用戶 set 銀兩=銀兩+" & money & " where id="&info(9))
 conn.close
 set conn=nothing
%>
<html>
<head>
<title>礦石收購場</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "新細明體"; font-size: 12px}
-->
</style>
</head>

<body bgcolor="#376D95">
<br>
<table width="72%" border="0" cellspacing="1" cellpadding="2" bgcolor="#999999" align="center">
  <tr align="center" bgcolor="#376D95"> 
    <td colspan="2" height="22"><font color="#FFCCCC"><b><%=info(0)%></b></font><b><font color="#FFFFFF">你把身上礦石全賣了，獲得銀兩</font><font color="#FFCCCC"><%=money%></font><font color="#FFFFFF">兩</font></b></td>
  </tr>
  <tr align="center" bgcolor="#376D95">
    <td colspan="2" height="26">
      <input type="submit" name="Submit" value="關閉" onclick="window.close()">
    </td>
  </tr>
</table>
</body>
</html>
