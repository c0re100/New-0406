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
<title>�� �F�������f-�A���F�[���j�h</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../css.css" rel=stylesheet>
</head>

<body bgcolor="#990000" background="../chat/bk.jpg" leftmargin="0" topmargin="0">
<table border="1" width="360" cellpadding="8"
cellspacing="8" background="../images/12.jpg">
  <tr bgcolor="#FFFFFF" align="center">
    <td background="../ff.jpg" bgcolor="#FFFFFF"> �Фj�L�J�Ӷ�g�H�U�H��------<font color="#FF3333">��������</font> 
      <form method="POST" action="casper.asp">
        <table width="254" align="center" height="194">
          <tr>
            <td height="126"> 
              <div align="center">
�m�W�G
<input type="text" name="name" size="10" maxlength="10">
<br>
<br>
�K�X�G
<input type="password" name="pass" size="10" maxlength="10">
<br>
 �{�ҡG 
<input type=text name=regjm1 size=5 maxlength="5" value="<%=regjm%>">
<br>
                ������ҳ��п�J�{�ҡG<font color="#FF0000"><%=regjm%></font> 
                <p>
              </div>
</td>
</tr>
<tr>
<td align="center"><input type=hidden name=regjm value="<%=regjm%>">
<input type="submit" name="submit" value="�_��">
<input type="button" value="����" onclick="window.close()"
name="button"></td>
</tr>
<tr>
<td>�F����:<br>
�@�w�n�O��o�Ӥ�l!!!�o�O�A���h����l,�ڥi�H���A����,�A�@�w�n�_��!!!<br>
</table>
</form>
</table>

</body>

</html> 