<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>
<title>會員數據庫管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#FFFFFF" background="0064.jpg">
<p align="center"><font size="+6" color="#FF0000">靈劍江湖數據庫</font></p>
<table width="572" border="1" align="center" cellspacing="0" cellpadding="0" bordercolor="#000000">
<tr>
<td>
<p align="center">&nbsp;</p>
<form method="POST" action="hyadminok.asp">
<div align="center">會員截止日期：
<select name="year" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:#000000 solid">
<option value="2004" selected>2004</option>
<option value="2005">2005</option>
<option value="2006">2006</option>
<option value="2007">2007</option>
<option value="2008">2008</option>
<option value="2009">2009</option>
<option value="2010">2010</option>
</select>
<select name="month" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
</select>
<select name="day" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
<option value="24">24</option>
<option value="25">25</option>
<option value="26">26</option>
<option value="27">27</option>
<option value="28">28</option>
<option value="29">29</option>
<option value="30">30</option>
<option value="31">31</option>
</select>
<br>
會員等級：
<select name="hygrade" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
<option value="1" selected>一級會員</option>
<option value="2">二級會員</option>
<option value="3">三級會員</option>
<option value="4">四級會員</option>
<option value="5">五級會員</option>
<option value="6">六級會員</option>
<option value="7">七級會員</option>
<option value="8">八級會員</option>
<option value="9">九級會員</option>
<option value="10">十級會員</option>
</select>
<br>
會員姓名：
<input type="text" name="search" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">會員方式：
<select name="hylx" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
<option value="1" selected>會員</option>
<option value="2">免費會員</option>
</select>
<br>
<input type="submit" value="確定新會員" name="B1" class="p9">
</div>
        <div align="center"><font color="#FF0000"><br>
          千萬不要拿站長設置會員，付費制會員請手動設置！</font><br>
<br>
</div>
<div align="center">
<center>
</center>
</div>
<div align="center">請選擇所要查姓名及等級設置！<br>
          需要注意：會員截止日期不能小於當前日期，會員在到期前5天提示！由於彙款時間問題，會員到期隻提示不扣級，扣級由站長親自修改！<br>
          不同等級的會員加的也不同的,如果要修改請手動改！<br>
</div>
</form>
      <p align="center">[站長可以通過此處來修改用戶的所以資料]<br>
        <a href="sendwupin.asp">物品發送</a> </p>
<p align="center"><a href="hyadmc.asp">現有會員列表</a> <a href="fine.asp">更新用戶資料</a></p>
</td>
</tr>
</table>
</body>
</html>
