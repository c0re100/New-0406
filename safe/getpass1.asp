<html>
<head>
<title>找回密碼</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="CHAT/READONLY/STYLE.CSS">
</head>

<body bgcolor="#990000" TOPMARGIN="0" LEFTMARGIN="0">
<p align="center"><font color="#FFFFFF"> &lt;&gt; 找回密碼 &lt;&gt;</font></p><table width="347" border="1" cellspacing="13" cellpadding="10" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#008080" BACKGROUND="file:///F|/%C1%E9%BD%A3%BD%AD%BA%FE%D7%DC%D5%BE%D5%BE%B5%E3/IMAGES/12.JPG" HEIGHT="192"> 
<form method="post" action="getpass.asp" name="getpass"> <tr> <td colspan="4" BACKGROUND="../ljimage/downbg.jpg">用 
戶 名 <input type="text" name="username" size="10" value="" maxlength="10"> <br> 
你在江湖裡大名<br> 找回密碼 <input type="password" name="getpass" size="10" maxlength="10"> 
<br> 注冊時您輸入的取回密碼項<br> 你的郵箱 <input type="text" name="email" size="20" maxlength="20"> 
<br> 你注冊的郵箱地址 <div align="center"> <font size="2"> <input type="submit" name="Submit" value="取回"> 
<input type=button value=關閉 onClick="window.close()" name="button"> </font></div></td></tr> 
</form></table><table width="357" border="1" align="center" cellpadding="0" cellspacing="0" bgcolor="#99CCFF" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
<tr> <td height="68"><font color="#000000">此功能用於您遺忘密碼後重新初始化你的密碼！初始化密碼為&quot;123456&quot; 
初始化密碼後，請您趕快更改您的密碼. &quot;找回密碼&quot;不是您常用的密碼，它是您注冊用戶名時必要的，隻有知道用戶名、找回密碼和您的郵箱，您纔可重新初始化您的密碼！</font></td></tr> 
</table><p align="center">&nbsp;</p><p align="center"><font color="#FFFFFF" size="2"><%Response.Write "序列號：<font color=blue>" & Application("hxf_c_sn") & "</font>，授權給：<font color=blue>" & Application("hxf_c_user") & "</font><br><font color=999999>" & Application("hxf_c_copyright") & "</font>"%> 
</font> </p>
</body>
</html>