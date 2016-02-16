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
<title>懺悔崖</title>
<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"><img src="scene_03b.jpg" width="470" height="225"> </div>
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center" bgcolor="#006699">
  <tr> 
<td height="81" valign="top"> 
      <div align="center"><b><%=info(0)%>來到懺悔崖</b></div>
      <form method=POST action='gryok.asp'> 
<table width="300" align="center"> <tr> <td> <tr> 
            <td align=center>我花多少錢贖罪： 
              <select name=money size=1> 
<option value="1000" selected>1000</option> <option value="10000">10000</option> 
<option value="100000">100000</option> <option value="1000000">1000000 <option value="10000000">10000000</option> 
</select> </td></tr> <tr> <td  align=center> 
              <input type=submit value=給我個機會吧 name="submit"> 
</td></tr> <tr> <td valign="top" height="8" > <div align="center"><br>
              </div>
            </td></tr> 
<tr> <td valign="top" > 
              <p>在江湖裡壞事做盡，到頭來深思其過，纔發覺以前做了很多錯事，現在後悔還來得及，花錢消災吧！</p>
            </td></tr> 
</table></form></td></tr> </table>
</body>
</html>



