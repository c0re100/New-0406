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
<title>用戶數據庫管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#FFFFFF" topmargin="0" background="0064.jpg">
<p align="center"><font size="+6" color="#FF0000">靈劍江湖數據庫</font></p>
<p align="left">管理說明：<br>
  <font color="yellow"><br>
  </font>現在有一些字段沒有有：如知質、知質加、魅力、魅力加、會員費等這些是為了以後功能升級時使用的，是一些備用的。<font color="yellow"><br>
  </font><font color="#0000FF">會員時間是截止時間，會員等級為：1,2,3,4如果其值為0，則表示不是會員！</font><br>
  武功加，內力加，體力加，攻擊加等。這一些值是不要亂改的，這些是他們在江湖上打坐、閉目得到的這些值不要改動！如果在管理上有什麼不明白的問題，請找靈劍江湖 
  電郵:Seven.s-888@yahoo.com.tw</p>
<div align="center">用戶資料修改程序 </div>
<form method="POST" action="showuser.asp">
  <div align="center"> 
    <select name="sjcz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="姓名" selected>姓名</option>
      <option value="ID">ID</option>
      <option value="等級">等級</option>
      <option value="配偶">配偶</option>
      <option value="信箱">信箱</option>
      <option value="Oicq">OIcq</option>
    </select>
    <input type="text" name="search" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="10">
    <input type="submit" value="查詢" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  <div align="center">ID查找一定要為數字！<br>
  </div>
  </form>
<form method="POST" action="laren.asp">
  <div align="center">拉人查看程序！<br>
    <input type="text" name="larenseek" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="submit" value="查詢" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空">
  </div>
  <div align="center"></div>
</form>
<form method="POST" action="pass.asp">
  <div align="center">密碼修正程序<br>輸入人名將他的密碼改成：123456<br>
    <input type="text" name="cpass" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"  maxlength="10">
    <input type="submit" value="確定" name="B12" class="p9">
    <input type="reset" name="Submit2" value="清空">
    <br>
  </div>
  </form>
<form method="POST" action="jiaji.asp">
  <div align="center">玩家戰鬥級加級程序<br>
    輸入人名和所要加的級數<br>
    <input type="text" name="cpass2" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="text" name="cpass22" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="submit" value="確定" name="B122" class="p9">
    <input type="reset" name="Submit22" value="清空">
    <br>
  </div>
</form>
<form method="POST" action="jieji.asp">
  <div align="center">玩家戰鬥級減級程序<br>
    輸入人名和所要減的級數<br>
    <input type="text" name="cpass23" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="text" name="cpass222" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="10">
    <input type="submit" value="確定" name="B1222" class="p9">
    <input type="reset" name="Submit222" value="清空">
    <br>
  </div>
</form>
</body>
</html>
 