<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)=""  then
Response.Redirect "../error.asp?id=440"
else
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


dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "開什麼玩笑？您下這麼少的注賭個什麼勁？"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "開什麼玩笑？您賭這麼多想叫我破產啊！"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="Select  銀兩,魅力 from 用戶 where id="&info(9)
set rs=conn.Execute(sql)
nowcash=rs("銀兩")
if betcash>nowcash then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "有沒有搞錯？您老荷包裡哪有這麼多銀子？（嘿嘿嘿，快去掙點去吧，窮鬼！）"
location.href = "javascript:history.back()"
</script>
<%
else

nowmeili=rs("魅力")
if nowmeili<10 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "大俠，您目前的魅力低於10了，面子上看不過去呀，以後再來吧！"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<html>
<head>
<title>江湖賭館 - 賭大小</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>

<body text="#000000" background="../images/8.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
<p><font size="3" class="c" color="#000000"><b>賭 場 - 賭 大 小<br>
</b></font><font size="2" class="c"><b><font color="#FF0033">你押的是</font><font color="#CC0000">
<%if request.form("select")="big" then%>
</font><font color="#CC0000"> </font></b></font><img src="../jhimg/bbs/big.gif" width="46" height="40"><font size="2" class="c"><b><font color="#CC0000">
<%else%>
<img src="../jhimg/bbs/small.gif" width="46" height="40"></font><font size="2" class="c" color="#CC0000">
<%end if%>
</font></b></font>
<%if (m1+m2+m3)>9 then%>
</p>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" bordercolor="#000000">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>結果:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>點，大</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="32" height="31"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2" color="#FFFFFF">
<%if request.form("select")="small" then%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>莊家：</b></font></td>
<td colspan="2"><font size="2">呵呵還要來嗎？ 有賭未必輸哦~！</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">再來！為什麼不？，我呸！我就不相信這麼倒霉~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">算了算了，認倒霉吧，留著點銀子救命吧~！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%
conn.execute("Update 用戶 set 銀兩=銀兩-"&betcash&",魅力=魅力-1 where id="&info(9))
%>
</font><font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%> </b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> 兩</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>莊家：</b></font></td>
<td colspan="2"><font size="2">厲害呀~！客官您還要來嗎？ </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">當然再來！我今天運氣不錯~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
            <td colspan="2"><font size="2"><a href="../welcome.asp">見好就收，你以為我傻？嘿嘿！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update 用戶 set 銀兩=銀兩+"&betcash&",魅力=魅力-1 where id="&info(9))
%>
</font><font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font> </td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3" height="2"><a href="betindex.asp"><font size="2">賭</font><font size="2">場首頁</font></a></td>
</tr>
</table>
<font size="3"><%else%></font>
<table border="1" cellspacing="0" cellpadding="3" align="center" width="400" bordercolordark="#FFFFFF" height="160">
<tr bgcolor="#336633">
<td colspan="3" align="center"><font class="c" size="2" color="#FFFFFF"><b>結果:</b></font><font size="2" color="#FFFFFF"><b><%=(m1+m2+m3)%>點，小</b></font></td>
</tr>
<tr>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif"></font></td>
<td width="33%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"></font></td>
<td width="34%" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2" color="#FFFFFF">
<%if request.form("select")="big" then%>
</font> <font size="2" color="#FFFFFF"><b><%=betcash%></b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right"><font size="2" color="#000099"><b>莊家：</b></font></td>
<td colspan="2"><font size="2">呵呵還要來嗎？ 有賭未為輸哦~！</font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">再來！為什麼不？，我呸！我就不相信這麼倒霉~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">算了算了，認倒霉吧，留著點銀子救命吧~！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update 用戶 set 銀兩=銀兩-"&betcash&",魅力=魅力-1 where id="&info(9))
%>
</font> <font size="2" color="#FFFFFF"><b><%=(nowcash-betcash)%></b><%=(nowmeili-1)%>
<%else%>
</font><font size="2" color="#FFFFFF"><b><%=betcash%> 兩</b></font>
<table width="100%" border="1" cellspacing="0" cellpadding="5" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC">
<tr>
<td align="right"><font size="2" color="#000099"><b>莊家：</b></font></td>
<td colspan="2"><font size="2">厲害呀~！客官您還要來嗎？ </font></td>
</tr>
<tr>
<td align="right"><font size="2" color="#CC0000"><b>我說：</b></font></td>
<td colspan="2"><font size="2"><a href="b&amp;s.asp">當然再來！我今天運氣不錯~！</a></font></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td colspan="2"><font size="2"><a href="../jh.asp">見好就收，你以為我傻？嘿嘿！</a></font></td>
</tr>
</table>
<font size="2" color="#FFFFFF">
<%conn.execute("Update 用戶 set 銀兩=銀兩+"&betcash&",魅力=魅力-1 where id="&info(9))
%>
</font> <font size="2" color="#FFFFFF"><b><%=(betcash+nowcash)%></b><%=(nowmeili-1)%></font><font size="2" color="#FFFFFF">
<%end if%>
</font>
<p>&nbsp;</p>
</td>
</tr>
<tr bgcolor="#FFCCCC">
<td align="right" colspan="3"><a href="betindex.asp"><font size="2">賭場首頁</font></a></td>
</tr>
</table>
<%end if%>
</div>
</body>
</html>
<%rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
set username=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
end if
%>
