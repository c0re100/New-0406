<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%><%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>嬰兒喂養</title>
<style type="text/css">
<!--
body {  font-size: 9pt}
a:link {  text-decoration: none}
a:hover {  text-decoration: underline}
-->
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="bk.jpg" bgproperties="fixed">
<div align="center"> 
  <p>此項功能保留，有待開發！<br>
    <br>
  </p>
  <table width="264" border="0">
    <tr> 
      <td colspan="4">
        <div align="center"><img src="file:///I|/ljjhduo/HCJS/JHJS/001/701.gif" width="69" height="83"></div>
      </td>
    </tr>
    <tr> 
      <td width="65"><img src="file:///I|/ljjhduo/HCJS/JHJS/001/907.gif" width="21" height="30"></td>
      <td colspan="2"><img src="file:///I|/ljjhduo/HCJS/JHJS/001/402.gif" width="88" height="99"></td>
      <td width="86"><img src="file:///I|/ljjhduo/HCJS/JHJS/001/008.gif" width="73" height="46"></td>
    </tr>
    <tr>
      <td width="65">&nbsp;</td>
      <td colspan="2"><img src="file:///I|/ljjhduo/HCJS/JHJS/001/203.gif" width="72" height="55"></td>
      <td width="86">&nbsp;</td>
    </tr>
  </table>
  <p><font color="#FFFFFF">嬰兒</font><font size="3"><br>
    <br>
    </font><font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> </font></span></font><font size="3" color="#000000"> 
    喂 養</font><font color="#FFFFFF"><font size="3">  </font></font><font size="3"><br>
    </font><font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> <br>
     </font></span></font><font size="3" color="#000000"> 擁 抱 </font><font color="#FFFFFF"><font size="3"> </font></font><br>
    <br>
    <font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> </font></span></font><font size="3" color="#000000"> 
    遊 戲 </font><font color="#FFFFFF"><font size="3"> </font></font></p>
<p> <font color="#FFFFFF"><span style='font-size:9pt'><font size="3"> </font></span></font><font size="3" color="#000000">
學 習</font><font color="#FFFFFF"><font size="3">  </font></font></p>
<p>這是我們將要推出的新功能，成家的夫妻一個月後可以選擇嬰兒功能，將由計算機給夫婦設計一個嬰兒由夫妻共同喂養，如果離婚，嬰兒將會進入孤兒院，具體細節還有侍開發！</p>
</div>
</body>
</html>
 