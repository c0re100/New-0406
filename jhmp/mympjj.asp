<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>門派基金</title>
<body bgcolor="#660000" background="../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<table width="97" border="0" align="center" cellpadding="0" cellspacing="0" background="../chat/bk.jpg">
  <tr> 
<td height="81" valign="top"> 
      <div align="center">歡迎<font color="#000000"><font color=#FFFFCC><%=info(0)%></font></font>光臨門派基金庫</div>
      <form method=POST action='putmpjj.asp'> 
<table width="300" align="center"> <tr> <td> <tr> <td align=center>請選擇所以存入的錢數： 
<select name=money size=1 style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
                <option value="1000" selected>1000</option>
                <option value="10000">10000</option>
                <option value="100000">100000</option>
                <option value="1000000">1000000</option>
                <option value="10000000">10000000</option>
              </select> </td></tr> <tr> <td  align=center> <input type=submit value=好了就這樣 name="submit"> 
</td></tr> <tr> <td valign="top" height="8" > <div align="center"><br> <br> 門派基金</div></td></tr> 
<tr> <td valign="top" > <p>門派基金，是一項取之於民，用之於民的政策，所取得金額存放於本門派金庫中，由執法長老分配給所需要的人。您對本派的付出我會有應有的回報的。對於本門支持最大或從本門取錢最多的人我們都會有記錄排行的，請到江湖排行中查詢！</p></td></tr> 
</table></form></td></tr> </table>
</body>
</html>
