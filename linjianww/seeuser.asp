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
<p align="left">請輸入條件查詢用戶：<br>
  如：身份='掌門' 查詢身份為掌門的用戶。<br>
  再如： 銀兩&gt;10000 and 武功&gt;10000 查看銀兩大於10000並且武功大於10000的用戶。<br>
  在查詢中可以使用：&quot;and&quot; &quot;or&quot; &quot;&gt;&quot; &quot;&lt;&quot; &quot;&lt;&gt;&quot; 
  &quot;&gt;=&quot; &quot;&lt;=&quot; &quot;=&quot; 關繫量！<br>
  在查詢物品時：查詢字段的值無效，如輸入：擁有者='江湖總站'<br>
  <font color="#0000FF">查詢字段：此值為單一值,如：內力、體力等。以下為錯誤：內力 and 體力 或 內力,體力等</font></p>
<div align="center">用戶資料修改程序 </div>
<form method="POST" action="seeuserok.asp">
  <div align="center"><font color="#FFFFFF">
    <select name="seekfs" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="0" selected>查詢用戶數據庫</option>
      <option value="1">查詢物品數據庫</option>
      <option value="2">查詢修煉物品庫</option>
    </select>
    </font><br>
    <br>
    請輸入查詢條件： 
    <input type="text" name="tiaojian" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="50" maxlength="250">
    <br>
    輸入將查詢字段： 
    <input type="text" name="show" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10" maxlength="12">
    <br>
    <input type="submit" value="查詢" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  </form>
</body>
</html> 