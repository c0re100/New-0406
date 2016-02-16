<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")="" then
response.redirect "index.asp"
end if
randomize
m=int(6*rnd+1)
if m>3 then
if request.form("select")="big" then
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(5 * Rnd + 1)
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(7 * Rnd + 1)
if m2>6 then m3=6 end if
end if
else
Randomize
m1 = Int(6 * Rnd + 1)
Randomize
m3 = Int(6 * Rnd + 1)
Randomize
m2 = Int(6 * Rnd + 1)
end if



if request.form("select")="big" then
if m1+m2+m3>9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3>=10
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
else
if m1+m2+m3<9 then
mm=int(6*rnd+1)
if mm=1 or mm=2 or mm=3 or mm=4 then
do while m1+m2+m3<9
m1 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m2 = Int(6 * Rnd + 1)
loop
end if
end if
end if
dim betcash,nowcash,username
username=ljjh_nickname
Session("iq")=""
%>

<head>
<title>結果</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet>
</head>

<body bgcolor="#000000" text="#FFFFFF" background="images/bg.gif">
<font color="#FFFFFF"><br>
</font>
<div align="center">
<p><font size="3" class="c" color="#FFFFFF"><b><font color="#000000">最後結果</font></b></font></p>
<p><font size="2" class="c"><b><font color="#FF0033">你押的是</font><br>
<font color="#CC0000">
<%if request.form("select")="big" then%></font><font color="#CC0000">
</font></b></font><img src="big.gif" width="46" height="40"><font size="2" class="c"><b><font color="#CC0000"><%else%><img src="small.gif" width="46" height="40"></font><font size="2" class="c" color="#CC0000"><%end if%></font></b></font><%if (m1+m2+m3)>9 then%>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>結果:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>點，大</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="images/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="images/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="images/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF"><br>
<%if request.form("select")="small" then%></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>司徒輸光：</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="javascript:self.close()"><font color="#FFFFFF">我贏了，你還是先回去練練賭術吧。</font></a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2" width="78%"><font size="2">真倒霉!!</font></td>
</tr>
</table>
<font size="2" color="#FFFFFF"><%else%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>司徒輸光：</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="go8.asp"><font color="#FFFFFF">我輸了，你按這裡繼續闖關吧。</font>
</a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2" width="78%"><font size="2">當然！我今天運氣不錯~！</font></td>
</tr>
</table>
<%
Session("iq")="ok"
%>
<%end if%>
<p>&nbsp;</p>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="9">&nbsp;</td>
</tr>
</table>
<font size="3"><%else%></font>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>結果:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>點，小</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="images/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="images/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="images/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF"><br>
<%if request.form("select")="big" then%></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>司徒輸光：</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="javascript:self.close()"><font color="#FFFFFF">我贏了，你還是先回去練練賭術吧。</font></a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2" width="78%"><font size="2">真倒霉!!</font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%else%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" width="22%"><font size="2" color="#000099"><b>司徒輸光：</b></font></td>
<td colspan="2" width="78%"><font size="2"><a href="go8.asp"><font color="#FFFFFF">我輸了，你按這裡繼續闖關吧。</font>
</a></font></td>
</tr>
<tr>
<td align="right" width="22%"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2" width="78%"><font size="2">當然！我今天運氣不錯~！</font></td>
</tr>
</table>
<%
Session("iq")="ok"
%>
<%end if%>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="9">&nbsp;</td>
</tr>
</table>
<%end if%></div>

</body>
</html>