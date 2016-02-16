<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
dim betcash,nowcash
nowcash=0
betcash=0
betcash=clng(request.form("splosh"))
if betcash<10 then
%>
<script language=vbscript>
MsgBox "開什麼玩笑？您下這這麼少賭個什麼勁呀？"
location.href = "javascript:history.back()"
</script>
<%
elseif betcash>3000 then
%>
<script language=vbscript>
MsgBox "開什麼玩笑？您賭這麼多想叫我破產啊!"
location.href = "javascript:history.back()"
</script>
<%
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="Select 銀兩,魅力 From 用戶 Where id="&info(9)
set rs=conn.Execute(sql)

Randomize
Randomize
m1 = Int(6 * Rnd + 1)
m4 = Int(6 * Rnd + 1)

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

Randomize
m2 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m3 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)

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

m=Second(time())
m=right(m,1)
if m="0" or m="1" or m="6" or m="7" or m="8" then
do while m1+m2+m3<=m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
else
do while m1+m2+m3>=m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
end if
if m1+m2+m3<m4+m5+m6 then
mm=int(rnd*6+1)
if mm=1 or mm=2   then
do while m1+m2+m3<m4+m5+m6

m4 = Int(6 * Rnd + 1)
m5 = Int(6 * Rnd + 1)
m6 = Int(6 * Rnd + 1)
loop
end if
end if




%>

<html>
<head>

<title>賭骰子</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
body,table {font-size: 9pt; font-family: 新細明體}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>

</head>

<body text="#000000" background="../images/8.jpg" topmargin="0">
<font color="#FFFFFF"><br>
</font>
<div align="center">
<p><font size="2" class="c" color="#FFFFFF"><font size="3"><b><font color="#000000">賭
場 - 賭 骰 子</font></b></font></font></p>


<table border="1" cellspacing="0" cellpadding="3" align="center" bordercolordark="#FFFFFF" bordercolorlight="#000000" bgcolor="#CCCCCC">
<tr>
<td bgcolor="#336633" colspan="3" height="23"><font size="2" class="c"><b>&nbsp;&nbsp;<font color="#FFFFFF">賭
局 結 果</font></b></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2">□-莊家骰子：<%=(m1+m2+m3)%> 點 </font></td>
</tr>
<tr>
<td width="100" align="center"><font size="2"><img src="../jhimg/bbs/<%=m1%>.gif" width="31" height="32"></font></td>
<td width="100" align="center"><font size="2"><img src="../jhimg/bbs/<%=m2%>.gif"width="31" height="32"></font></td>
<td width="139" align="center"><font size="2"><img src="../jhimg/bbs/<%=m3%>.gif" width="32" height="32"></font></td>
</tr>
<tr>
<td colspan="3" align="center"><font size="2">□-你的骰子：<%=(m4+m5+m6)%> 點 </font></td>
</tr>
<tr>
<td width="100" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m4%>.gif"></font></td>
<td width="100" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m5%>.gif"></font></td>
<td width="139" align="center" height="39"><font size="2"><img src="../jhimg/bbs/<%=m6%>.gif"></font></td>
</tr>
<tr>
<td colspan="3" align="center"> <font size="2">
<%if (m1+m2+m3)>=(m4+m5+m6) then%>
</font>
<table width="100%" border="1" cellspacing="0" cellpadding="0" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td height="49" width="14%">&nbsp;</td>
<td colspan="2" height="49" width="86%">
<p><font size="2"><img src="../jhimg/bbs/face13.gif" width="20" height="20" align="absmiddle">
真倒霉，輸了： <b><font color="#CC0000"><%=betcash%> </font>兩</b></font></p>
</td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b> 莊家：</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#000099"><img src="../jhimg/bbs/face18.gif" width="20" height="20" align="absmiddle"></font></b>
呵呵~還要來嗎？ 有賭未必輸哦~！</font></td>
</tr>
<tr>
<td align="right" rowspan="2" width="14%"><font size="2"><b> 我要：</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#CC0000"><img src="../jhimg/bbs/face8.gif" width="20" height="20" align="absmiddle"></font></b>
<a href="dice.asp">再來！為什麼不？，我呸！我就不相信這麼倒霉~！</a></font></td>
</tr>
<tr>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face11.gif" width="20" height="20" align="absmiddle">
<a href="../jh.asp">算了算了，認倒霉吧，留著點銀子救命吧~！</a></font></td>
</tr>
</table>
<font size="2"><%
conn.execute("Update 用戶 set 銀兩=銀兩-"&betcash&",魅力=魅力-1 where id="&info(9))
%> </font>
<hr size="1" width="250">
<font size="2"> 你現在有銀子：<font color="#CC0000"><b><font color="#CC0000"><%=(nowcash-betcash)%></font>
</b></font>兩 魅力：<%=(nowmeili-1)%><%else%> </font>
<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td align="right" height="50" width="14%">&nbsp;</td>
<td colspan="2" height="50" width="86%">
<p><font size="2"><img src="../jhimg/bbs/face14.gif" width="20" height="20" align="absmiddle">
嘻嘻~，贏了：<b><font color="#CC0000"><%=betcash%> </font>兩</b></font></p>
</td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b> 莊家：</b></font></td>
<td colspan="2" width="86%"><font size="2"><b><font color="#000099"><img src="../jhimg/bbs/face17.gif" width="20" height="20" align="absmiddle"></font></b>
厲害呀~！客官您還要來嗎？ </font></td>
</tr>
<tr>
<td align="right" width="14%"><font size="2"><b>我要：</b></font></td>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face10.gif" width="20" height="20" align="absmiddle">
<a href="dice.asp">當然再來！我今天運氣不錯~！</a></font></td>
</tr>
<tr>
<td align="right" width="14%">&nbsp;</td>
<td colspan="2" width="86%"><font size="2"><img src="../jhimg/bbs/face2.gif" width="20" height="20" align="absmiddle">
<a href="../jh.asp">見好就收，你以為我傻？嘿嘿！</a></font></td>
</tr>
</table>
<font size="2"><%conn.execute("Update 用戶 set 銀兩=銀兩+"&betcash&",魅力=魅力-1 where id="&info(9))%>
</font>
<hr size="1" width="250">
<font size="2"> 你現在有銀子：<b><font color="#CC0000"><b><%=(betcash+nowcash)%></b>
</font>兩 </b>魅力: <%=(nowmeili-1)%><%end if%> </font>
<hr size="1" width="250">
</td>
</tr>
<tr>
<td align="right" colspan="3"><font size="2"><a href="betindex.asp">賭場首頁</a></font></td>
</tr>
</table>
</div>
</body>
</html>
<%rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
set betcash=nothing
set nowcash=nothing
end if
end if
end if
%>

<html><script language="JavaScript">                                                                  </script></html>