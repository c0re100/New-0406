<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 二婚 FROM 用戶 where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "error.asp?id=453"
	response.end
end if
if rs("二婚")="" or rs("二婚")="無" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&info(0)&"你還沒有二婚呢，來作什麼！]');history.go(-1);</script>"
	response.end
end if
%>
<html>
<head>
<title>自創夫妻合體技</title>
<style></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#FFFFFF" bgcolor="#660000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<div align="center"> <b>夫 妻 合 體 技 設 置</b><font color="#000000"><br>
<br>
</font> </div>
<table cellpadding="0" cellspacing="1" border="1" align="center" width="597" bgcolor="#00CCFF" bordercolor="#000000">
<tr>
<td height="11" width="98">
<div align="center"> <font color="#000000" size="2">招 式 名</font> </div>
</td>
<td height="11" width="124">
<div align="center"> <font color="#000000" size="2">所 用 內 力</font> </div>
</td>
<td height="11" width="367">
<div align="center"> <font color="#000000" size="2">操 作</font> </div>
</td>
</tr>
<%
rs.close
rs.open "SELECT 合體技,內力 FROM 合體技 where 姓名男='" & info(0) & "' or 姓名女='" & info(0) & "'"
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
<form method=POST action='stunt1.asp?a=m&wg=<%=rs("合體技")%>'>
<tr>
<td height="2" width="98">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="wg1" size="10"
value="<%=rs("合體技")%>" maxlength="8">
</font> </div>
</td>
<td height="2" width="124">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="nl" size="10"
value="<%=rs("內力")%>" maxlength="8">
</font> </div>
</td>
<td height="2" width="367">
<div align="center"> <font color="#000000" size="2">姓名：
<input type="text" name="name"
size="10" maxlength="10">
密碼：
<input type="password" name="pass"
size="10" maxlength="20">
<input type="submit" value="修改"
name="submit">
</font> </div>
</td>
</tr>
</form>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<10 then
%>
<form method=POST action='stunt1.asp?a=n'>
<tr>
<td width="98" height="2">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="wg1" size="10"
maxlength="8">
</font> </div>
</td>
<td width="124" height="2">
<div align="center"> <font color="#000000" size="2">
<input type="text" name="nl" size="10"
maxlength="8">
</font> </div>
</td>
<td width="367" height="2">
<div align="center"> <font color="#000000" size="2">姓名：
<input type="text" name="name"
size="10" maxlength="10">
密碼：
<input type="password" name="pass"
size="10" maxlength="20">
<input type="submit" value="添加"
name="submit">
</font> </div>
</td>
</tr>
</form>
<%end if%>
</table>
<p class="p1" align="center"><font size="2">[注：只有二婚大俠或俠女纔可以自創，每創建一次收取10000兩，創建後雙方都可使用，離婚後合體技失效！]<br>
  <br>
  [殺傷力=男方殺傷力+女方殺傷力]</font><font color="#000000" size="2"><br>
  <br>
  <br>
  </font></p>
</body>
</html> 