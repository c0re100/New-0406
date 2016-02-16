<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>溫泉浴中心</title>
<link rel="StyleSheet" href="../../css.css" title="Contemporary">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body background="../../ff.jpg" leftmargin="0" topmargin="0">
<div align="center">
<center>
    <table width="348" border="0" cellpadding="0" cellspacing="0" bgcolor="#003366">
      <tr bgcolor="#FFFFFF"> 
        <td></td>
      </tr>
      <tr> 
        <td> <form method="POST" action="CHECKSEX.ASP">
            <br>
            <br>
            <table width="100%" cellspacing="1">
              <tr> 
                <td class="p2" width="100%"> <font color="#CCCCCC">歡 迎 進 入 靈 劍 
                  江 湖 龍 泉 洗 浴 中 心&nbsp;</font> </td>
              </tr>
              <tr > 
                <td class="p3" width="100%" height="29"> <font color="#FFFFFF">進入洗浴後就不能領工資！！體力1000 
                  </font> </td>
              </tr>
              <tr> 
                <td class="p2" width="100%" height="50"> <input type="radio" value="男"
name="SEX" checked> <font color="#FFFFFF"> <font color="#CCCCCC">男賓部&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input type="radio" value="女" name="SEX">
                  女賓部</font></font> </td>
              </tr>
              <tr> 
                <td class="p3" width="100%"> <p>&nbsp;</p>
                  <p> 
                    <input type="submit" value="進入洗浴"
name="B1">
                  </p></td>
              </tr>
            </table>
          </form></td>
      </tr>
    </table>
</center>
</div>

</body>

</html>
