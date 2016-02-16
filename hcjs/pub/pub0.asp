<%
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖客棧</title></head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="#996699">
<p align="center" style="font-size:16;color:yellow"><font color="#000000">歡迎<%=info(0)%>光臨本客棧休息 
</font> <table border="1" align="center" width="399" cellpadding="10"
cellspacing="13" BACKGROUND="../../IMAGES/BJ31.JPG"> <tr> <td bgcolor="#BEE0FC" height="230" BACKGROUND="../../ljimage/downbg.jpg"> 
<table width="358"> <tr> <td> <form method="POST" action="pub1.asp"> <tr> <td align="center">你想要休息： 
<select name="time" size="1"> <option value="0">0小時 <option value="1" selected>1小時 
<option value="2">2小時 <option value="3">3小時 <option value="4">4小時 <option value="5">5小時 
<option value="6">6小時 <option value="7">7小時 <option value="8">8小時 <option value="9">9小時 
<option value="10">10小時 <option value="11">11小時 <option value="12">12小時 <option value="13">13小時 
<option value="14">14小時 <option value="15">15小時 <option value="16">16小時 <option value="17">17小時 
<option value="18">18小時 <option value="19">19小時 <option value="20">20小時 <option value="21">21小時 
<option value="22">22小時 <option value="23">23小時 </select></td></tr> <tr> <td colspan="2" align="center"><input type="submit" value="確定"> 
<INPUT TYPE="button" VALUE="關閉" ONCLICK="javascript:window.close();" NAME="button"></td></tr> 
<tr> <td colspan="2" style="font-size:9pt"> <hr> 1、在客棧中休息可以保護帳號，並可以增加內力和生命值；<br> 
2、每1小時收服務費10兩！增加生命10，內力10；<br> 3、請準確計算你下次使用帳號的時間！</td></tr> </form> </table></table>
</body>

</html>
