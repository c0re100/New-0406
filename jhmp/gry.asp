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
<title>孤兒院</title>
<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center"> <tr> 
<td height="81" valign="top"> 
      <div align="center"><b><%=info(0)%>光臨靈劍孤兒院</b></div>
      <form method=POST action='gryok.asp'> 
<table width="300" align="center"> <tr> <td> <tr> <td align=center>我的愛心： <select name=money size=1> 
<option value="1000" selected>1000</option> <option value="10000">10000</option> 
<option value="100000">100000</option> <option value="1000000">1000000 <option value="10000000">10000000</option> 
</select> </td></tr> <tr> <td  align=center> <input type=submit value=祝福這些孩子 name="submit"> 
</td></tr> <tr> <td valign="top" height="8" > <div align="center"><br> <br> 孤兒院簡介</div></td></tr> 
<tr> <td valign="top" > <p>江湖的戰爭、貧窮使得他們這些孩子失去了父母親，無助他的他們多麼希望得到我們最真誠的幫助，請獻出你的愛心，為世界所的孤兒捐獻一片愛心！(捐獻愛心可以使您德高望眾的1點道德=500兩！)</p></td></tr> 
</table></form></td></tr> </table>
</body>
</html>



