<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
id=request("id")
sql="SELECT * FROM 用戶 WHERE 體力<-1000 and 狀態='死' and 門派='官府' and ID=" & id
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess="此人未死或是自殺用不著你多心！"
else
sql="update 用戶 set 狀態=' ', 體力=0 where id=" & id
conn.execute sql
mess="成功地把ID" & id & "救活"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖版權所有</title></head>

<body background="../images/8.jpg">
<table border="1" align="center" width="350" cellpadding="10"
cellspacing="13" height="134" background="../images/12.jpg">
  <tr>
    <td background="../images/7.jpg"> 
      <table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="admin.asp">返回</a></p>
</td>
</tr>
</table>
    </td>
</tr>
</table>
<div align="center">
<br>
  <font color="#FFFFFF" size="2">★靈劍江湖版權所有★</font></div>

</body>
<html></html> 