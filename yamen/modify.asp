<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>修改密碼</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body background="../ff.jpg" LEFTMARGIN="0" TOPMARGIN="0">
<center> <table border=1 align=center width=326 cellpadding="5" cellspacing="8" background="../images/12.jpg" HEIGHT="350"> 
<tr bgcolor="#FFFFFF" align="center"> <td height="209" background="../ljimage/downbg.jpg"> 
<p>修 改 密 碼</p><table border="0" width="288" height="257"> <tr> <td> <form method=POST action='modifyok.asp'> 
<div align="center">姓 名： <input type=text name=name size=12 maxlength="10"> <br> 
原密碼： <input type=password name=oldpass size=12 maxlength="10"> <br> 新密碼： <input type=password name=pass size=12 maxlength="10"> 
<br> 確 認： <input type=password name=repass size=12 maxlength="10"> <br> 認 正： <input type=text name=regjm1 size=5 maxlength="5" value=<%=regjm%>> 
<br> 為防止朔雪請輸入認證：<font color="#FF0000"><%=regjm%></font> <br> <p> <input type=submit value=修改 name="submit"> 
<input type=button value=關閉 onClick="window.close()" name="button"> <input type=hidden name=regjm value="<%=regjm%>"> 
</p></div></form></td></tr> <tr> <td style='color:red;font-size:9pt'> <font color="#FF6600">說明：<br> 
<font color="#9900CC">(密碼長度不得超過10位）</font><br> 由於IE5會記錄輸入的密碼，<br> 請在網吧上網的朋友通過“清除歷史記錄”<br> 
來刪除記錄！以免帳號被盜用 </font></td></tr> </table></table></center>
</body>
</html>
 