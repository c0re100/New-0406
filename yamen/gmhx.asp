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
if info(0)<>"天之嬌子"   then Response.Redirect "../error.asp?id=439"
randomize timer
regjm=int(rnd*9998)+1
%>
<html>
<head>
<title>靈劍江湖 </title>
<style type="text/css">body, td     { font-size: 14 }
input        { font-size: 14; color: #000000 }
.p1          { font-size: 21pt; color: #ff0000 }
.p2          { font-size: 9pt; color: #00ee00 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body topmargin="0" bgcolor="#88AFD7">
<p class="p1" align="center"><b>改名換姓</b></p>
<table border="1" bgcolor="#FFCC99" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
    <td bgcolor="#88AFD7"> 
      <form method="POST" action="gmhxok.asp">
<table width="100%">
<tr>
            <td><font size="-1">姓名: 
              <input type="text" name="name" size="10" maxlength="10">
              </font> </td>
            <td><font size="-1">密碼: 
              <input type="password" name="pass" size="10" maxlength="10">
              </font> </td>
</tr>
<tr>
            <td height="2"><font size="-1">新名: 
              <input type="text" name="name1" size="10" maxlength="10">
              </font> </td>            
            <td height="2"> <font size="-1">認證: 
              <input type=text name=regjm1 size=5 maxlength="5">
              <br>
              防止朔雪輸入認證:<font color="#FF0000"><%=regjm%></font></font></td>
</tr>
<tr>
<td align="center" colspan="2"><input type=hidden name=regjm value="<%=regjm%>"><input type="submit" value="修改"
name="submit"> <input type="reset" value="取消" name="reset"></td>
</tr>
<tr>
            <td align="center" colspan="2"><b><font color="#FF0000">改名換姓：現金50萬兩</font></b></td>
</tr>
</table>
</form>
</td>
</tr>
</table>

</body>

</html>