<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("ljjh_jm")=session("ljjh_jm")+1
if session("ljjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>改名換姓</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body topmargin="0" bgcolor="#CC3300" text="#FFFFFF">
<p class="p1" align="center"><b><br>
  <font size="+6" color="#FFFF00">改名換姓</font></b></p>
<table border="1" bgcolor="#FFCC99" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
    <td bgcolor="#009933"> 
      <form method="POST" action="gmhxok.asp">
        <table width="100%">
          <tr> 
            <td><font size="-1">姓名: 
              <input type="text" name="name" size="10" maxlength="10" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
              </font> </td>
            <td><font size="-1">密碼: 
              <input type="password" name="pass" size="10" maxlength="10" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
              </font> </td>
          </tr>
          <tr> 
            <td height="2"><font size="-1">新名: 
              <input type="text" name="name1" size="10" maxlength="10" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
              </font> </td>
            <td height="2"> <font size="-1">認證: 
              <input type=text name=regjm1 size=5 maxlength="5" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
              <br>
              防止朔雪輸入認證:<font color="#FF0000"><%=regjm%></font></font></td>
          </tr>
          <tr> 
            <td align="center" colspan="2">
              <input type=hidden name=regjm value="<%=regjm%>">
              <input type="submit" value="修改"
name="submit">
              <input type="reset" value="取消" name="reset">
            </td>
          </tr>
          <tr> </tr>
        </table>
      </form>
    </td>
</tr>
</table>
</body>
</html>