<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
session("ljjh_jm")=session("ljjh_jm")+1
if session("ljjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
randomize timer
regjm=int(rnd*9998)+1
%><html>
<head>
<title>我不小心掉線</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body bgcolor="#990000" background="../chat/bk.jpg" leftmargin="0" topmargin="0">
<table border="1" width="360" cellpadding="10"
cellspacing="13" background="../images/12.jpg">
  <tr bgcolor="#FFFFFF">
    <td height="173" background="../ff.jpg" bgcolor="#FFFFFF"> 
      <div align="center">請大俠仔細填寫以下信息 </div>
<form method="POST" action="closeme.asp">
<table width="250" align="center">
<tr>
<td>
              <div align="center"><font color="#FFFFFF">
                <select name="chat">
                <%for i=0 to chatroomnum	
	chatroomxx=application("ljjh_chatroomname"&i)
%>
                  <option value="" selected><%=chatroomxx%></option>
               <% next%></select>
                </font> <br>
                請選擇你卡住的房間！<br>
                姓名： 
                <input type="text" name="name" size="10" maxlength="10">
<br>
                密碼： 
                <input type="password" name="pass" size="10" maxlength="10">
<br>
 認證： 
<input type=text name=regjm1 size=5 maxlength="5" value="<%=regjm%>">
<br>
                認證號：<font color="#FF0000"><%=regjm%></font> <br>
<p> 
</div>
</td>
</tr>
<tr>
<td align="center">
<input type=hidden name=regjm value="<%=regjm%>">
<input type="submit" name="submit" value="確定">
<input type="button" value="關閉" onclick="window.close()"
name="button"></td>
</tr>
<tr>
            <td style="color:red;font-size:9pt" height="18"> 
              <div align="center"> 輸入您的賬號及你賬號的密碼點擊確定即可！</div>
            
          </table>
</form>
</table>

</body>

</html>
 