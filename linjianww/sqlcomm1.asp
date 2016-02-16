<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
if info(2)<10  then Response.Redirect "manerr.asp?id=100"
%>
<html>
<head>
<title>sql指令系統</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<SCRIPT LANGUAGE="JavaScript">
function DoTitle(addTitle) {
var revisedTitle;
var currentTitle = document.PostTopic.sqlstr.value;
revisedTitle = addTitle;
document.PostTopic.sqlstr.value=revisedTitle;
document.PostTopic.sqlstr.focus();
return; }
</script>

<body bgcolor="#FFFFFF" background="../IMAGES/b115.jpg" text="#FFFFFF" topmargin="0">
<blockquote>
  <p align="center"> <br>
  </p>
</blockquote>
<%if sjjh_name=Application("sjjh_user") then%>
<form method="POST" action="sqlcommok.asp" name="PostTopic">
<div align="center">
<SELECT name=fs onchange=DoTitle(this.options[this.selectedIndex].value)> 
<OPTION selected value="">快速指令</OPTION> 
<OPTION value="update 用戶 set 銀兩=0,存款=0 where (銀兩+存款)>2000000000">[清除資產大於20億用戶資產為0]</OPTION>
<OPTION value="delete * from 操作記錄" >[刪除所有操作記錄]</OPTION>
<OPTION value="delete * from 人命" >[刪除所有人命記錄]</OPTION>
<OPTION value="delete * from 管理記錄" >[刪除所有管理記錄]</OPTION>
<OPTION value="delete * from 日記本" >[刪除所有日記本記錄]</OPTION>
<OPTION value="delete * from 申冤" >[刪除所有申冤記錄]</OPTION>
<OPTION value="delete * from 合體技" >[刪除所有合體技記錄]</OPTION>
<OPTION value="delete * from 離婚" >[刪除所有離婚記錄]</OPTION>
<OPTION value="delete * from 使用技能 where aszcc=0" >[刪除非會員打架招式]</OPTION>
<OPTION value="delete * from 使用技能 where 等級<100" >[刪除打架等級小於100]</OPTION>
<OPTION value="delete * from 物品 where 會員=0" >[刪除所有非會員物品]</OPTION>
<OPTION value="delete * from 交易市場 where 會員=0" >[刪除所有非會員保險櫃]</OPTION>
<OPTION value="delete * from 用戶 where 會員等級=0" >[刪除所有非會員]</OPTION>
</SELECT><br>
    <br>
    請輸入指令： 
    <input type="text" name="sqlstr" size="50" maxlength="280">
    <br>
    <input type="submit" value="執行" name="B1" class="p9">
    <input type="reset" name="Submit" value="清空">
  </div>
  </form>
 <%else%>
  <div align="center">您非授權站長，此項禁止使用！</div>
 <%end if%>
</body>
</html>