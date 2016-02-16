<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
randomize timer
regjm=int(rnd*9998)+1
%><html>
<head>
<title>◎ 閻王殿門口-你踫見了觀音大士</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../css.css" rel=stylesheet>
</head>

<body bgcolor="#990000" background="../chat/bk.jpg" leftmargin="0" topmargin="0">
<table border="1" width="360" cellpadding="8"
cellspacing="8" background="../images/12.jpg">
  <tr bgcolor="#FFFFFF" align="center">
    <td background="../ff.jpg" bgcolor="#FFFFFF"> 請大俠仔細填寫以下信息------<font color="#FF3333">報仇雪恨</font> 
      <form method="POST" action="casper.asp">
        <table width="254" align="center" height="194">
          <tr>
            <td height="126"> 
              <div align="center">
姓名：
<input type="text" name="name" size="10" maxlength="10">
<br>
<br>
密碼：
<input type="password" name="pass" size="10" maxlength="10">
<br>
 認證： 
<input type=text name=regjm1 size=5 maxlength="5" value="<%=regjm%>">
<br>
                為防止朔雪請輸入認證：<font color="#FF0000"><%=regjm%></font> 
                <p>
              </div>
</td>
</tr>
<tr>
<td align="center"><input type=hidden name=regjm value="<%=regjm%>">
<input type="submit" name="submit" value="復活">
<input type="button" value="關閉" onclick="window.close()"
name="button"></td>
</tr>
<tr>
<td>閻王說:<br>
一定要記住這個日子!!!這是你死去的日子,我可以讓你重生,你一定要復仇!!!<br>
</table>
</form>
</table>

</body>

</html> 