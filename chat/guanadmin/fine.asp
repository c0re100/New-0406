<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
if info(2)<8  then Response.Redirect "../../error.asp?id=425"
%>
<html>
<head>
<title>在線抓人特別管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../chat/READONLY/STYLE.CSS">
</head>

<body topmargin="0" bgcolor="#0066FF" text="#00FFFF">
<p align="center">在線抓人特別管理</p>
<p align="center">管理說明：<br>
  <font color="yellow"><br>
  現根據以下指示輸入名字、抓人原因點擊執行即可，如有不明，請找江湖總站聯繫</font></p>
<div align="center"></div>
<form method="POST" action="zhixin.asp">
  <div align="center">選擇管理菜單<br>
    輸入要抓人的名字、原因<br>
    <select name="ljgl">
      <option value="逮捕" selected>逮捕</option>
      <option value="坐牢">坐牢</option>
      <option value="監禁">監禁</option>
      <option value="寶物">寶物</option>
      <option value="炸彈">炸彈</option>
    </select>
    <input type="text" name="pass1" size="10" maxlength="10">
    <input type="text" name="pass11" size="10" maxlength="10">
    <input type="submit" value="執行" name="a1" class="p9">
    <input type="reset" name="Submit22" value="重填">
    <br>
  </div>
</form>
</body>
</html>
 