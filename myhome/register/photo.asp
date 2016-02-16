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

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖-頭像</title>
<link rel="StyleSheet" href="../new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#000000" text="#E18C59">

<div align="center">
<center>
<table border="0" width="776" cellspacing="0" cellpadding="0">
<tr>
<td width="100%"><IFRAME src="../topbanner.htm" width="776" height="60" marginwidth="0" marginheight="0" hspace="0" vspace="0" frameborder="0" scrolling="NO"></IFRAME></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td width="100%" class="p1" height="20">您當前的位置-更改個人頭像</td>
</tr>
</table>
</center>
</div>
<form method="POST" action="changeskin.asp">
<div align="center">
<center>
<table border="0" width="760" cellspacing="0" cellpadding="0">
<tr>
<td width="40"><img border="0" src="skin/01.gif"></td>
<td width="40"><img border="0" src="skin/2.gif"></td>
<td width="40"><img border="0" src="skin/3.gif"></td>
<td width="40"><img border="0" src="skin/4.gif"></td>
<td width="40"><img border="0" src="skin/5.gif"></td>
<td width="40"><img border="0" src="skin/6.gif"></td>
<td width="40"><img border="0" src="skin/7.gif"></td>
<td width="40"><img border="0" src="skin/8.gif"></td>
<td width="40"><img border="0" src="skin/9.gif"></td>
<td width="40"><img border="0" src="skin/10.gif"></td>
<td width="40"><img border="0" src="skin/11.gif"></td>
<td width="40"><img border="0" src="skin/12.gif"></td>
<td width="40"><img border="0" src="skin/13.gif"></td>
<td width="40"><img border="0" src="skin/14.gif"></td>
<td width="40"><img border="0" src="skin/15.gif"></td>
<td width="40"><img border="0" src="skin/16.gif"></td>
<td width="40"><img border="0" src="skin/17.gif"></td>
<td width="40"><img border="0" src="skin/18.gif"></td>
<td width="40"><img border="0" src="skin/19.gif"></td>
</tr>
<tr>
<td width="40" align="center"><input type="radio" value="1" name="skin"></td>
<td width="40" align="center"><input type="radio" value="2" name="skin"></td>
<td width="40" align="center"><input type="radio" value="3" name="skin"></td>
<td width="40" align="center"><input type="radio" value="4" name="skin"></td>
<td width="40" align="center"><input type="radio" value="5" name="skin"></td>
<td width="40" align="center"><input type="radio" value="6" name="skin"></td>
<td width="40" align="center"><input type="radio" value="7" name="skin"></td>
<td width="40" align="center"><input type="radio" value="8" name="skin"></td>
<td width="40" align="center"><input type="radio" value="9" name="skin"></td>
<td width="40" align="center"><input type="radio" value="10" name="skin"></td>
<td width="40" align="center"><input type="radio" value="11" name="skin"></td>
<td width="40" align="center"><input type="radio" value="12" name="skin"></td>
<td width="40" align="center"><input type="radio" value="13" name="skin"></td>
<td width="40" align="center"><input type="radio" value="14" name="skin"></td>
<td width="40" align="center"><input type="radio" value="15" name="skin"></td>
<td width="40" align="center"><input type="radio" value="16" name="skin"></td>
<td width="40" align="center"><input type="radio" value="17" name="skin"></td>
<td width="40" align="center"><input type="radio" value="18" name="skin"></td>
<td width="40" align="center"><input type="radio" value="19" name="skin"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="760" cellspacing="0" cellpadding="0">
<tr>
<td width="40"><img border="0" src="skin/20.gif"></td>
<td width="40"><img border="0" src="skin/21.gif"></td>
<td width="40"><img border="0" src="skin/22.gif"></td>
<td width="40"><img border="0" src="skin/23.gif"></td>
<td width="40"><img border="0" src="skin/24.gif"></td>
<td width="40"><img border="0" src="skin/25.gif"></td>
<td width="40"><img border="0" src="skin/26.gif"></td>
<td width="40"><img border="0" src="skin/27.gif"></td>
<td width="40"><img border="0" src="skin/28.gif"></td>
<td width="40"><img border="0" src="skin/29.gif"></td>
<td width="40"><img border="0" src="skin/30.gif"></td>
<td width="40"><img border="0" src="skin/31.gif"></td>
<td width="40"><img border="0" src="skin/32.gif"></td>
<td width="40"><img border="0" src="skin/33.gif"></td>
<td width="40"><img border="0" src="skin/34.gif"></td>
<td width="40"><img border="0" src="skin/35.gif"></td>
<td width="40"><img border="0" src="skin/36.gif"></td>
<td width="40"><img border="0" src="skin/37.gif"></td>
<td width="40"><img border="0" src="skin/38.gif"></td>
</tr>
<tr>
<td width="40" align="center"><input type="radio" value="20" name="skin"></td>
<td width="40" align="center"><input type="radio" value="21" name="skin"></td>
<td width="40" align="center"><input type="radio" value="22" name="skin"></td>
<td width="40" align="center"><input type="radio" value="23" name="skin"></td>
<td width="40" align="center"><input type="radio" value="24" name="skin"></td>
<td width="40" align="center"><input type="radio" value="25" name="skin"></td>
<td width="40" align="center"><input type="radio" value="26" name="skin"></td>
<td width="40" align="center"><input type="radio" value="27" name="skin"></td>
<td width="40" align="center"><input type="radio" value="28" name="skin"></td>
<td width="40" align="center"><input type="radio" value="29" name="skin"></td>
<td width="40" align="center"><input type="radio" value="30" name="skin"></td>
<td width="40" align="center"><input type="radio" value="31" name="skin"></td>
<td width="40" align="center"><input type="radio" value="32" name="skin"></td>
<td width="40" align="center"><input type="radio" value="33" name="skin"></td>
<td width="40" align="center"><input type="radio" value="34" name="skin"></td>
<td width="40" align="center"><input type="radio" value="35" name="skin"></td>
<td width="40" align="center"><input type="radio" value="36" name="skin"></td>
<td width="40" align="center"><input type="radio" value="37" name="skin"></td>
<td width="40" align="center"><input type="radio" value="38" name="skin"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="760" cellspacing="0" cellpadding="0">
<tr>
<td width="40"><img border="0" src="skin/39.gif"></td>
<td width="40"><img border="0" src="skin/40.gif"></td>
<td width="40"><img border="0" src="skin/41.gif"></td>
<td width="40"><img border="0" src="skin/42.gif"></td>
<td width="40"><img border="0" src="skin/43.gif"></td>
<td width="40"><img border="0" src="skin/44.gif"></td>
<td width="40"><img border="0" src="skin/45.gif"></td>
<td width="40"><img border="0" src="skin/46.gif"></td>
<td width="40"><img border="0" src="skin/47.gif"></td>
<td width="40"><img border="0" src="skin/48.gif"></td>
<td width="40"><img border="0" src="skin/49.gif"></td>
<td width="40"><img border="0" src="skin/50.gif"></td>
<td width="40"><img border="0" src="skin/51.gif"></td>
<td width="40"><img border="0" src="skin/52.gif"></td>
<td width="40"><img border="0" src="skin/53.gif"></td>
<td width="40"><img border="0" src="skin/54.gif"></td>
<td width="40"><img border="0" src="skin/55.gif"></td>
<td width="40"><img border="0" src="skin/56.gif"></td>
<td width="40"><img border="0" src="skin/57.gif"></td>
</tr>
<tr>
<td width="40" align="center"><input type="radio" value="39" name="skin"></td>
<td width="40" align="center"><input type="radio" value="40" name="skin"></td>
<td width="40" align="center"><input type="radio" value="41" name="skin"></td>
<td width="40" align="center"><input type="radio" value="42" name="skin"></td>
<td width="40" align="center"><input type="radio" value="43" name="skin"></td>
<td width="40" align="center"><input type="radio" value="44" name="skin"></td>
<td width="40" align="center"><input type="radio" value="45" name="skin"></td>
<td width="40" align="center"><input type="radio" value="46" name="skin"></td>
<td width="40" align="center"><input type="radio" value="47" name="skin"></td>
<td width="40" align="center"><input type="radio" value="48" name="skin"></td>
<td width="40" align="center"><input type="radio" value="49" name="skin"></td>
<td width="40" align="center"><input type="radio" value="50" name="skin"></td>
<td width="40" align="center"><input type="radio" value="51" name="skin"></td>
<td width="40" align="center"><input type="radio" value="52" name="skin"></td>
<td width="40" align="center"><input type="radio" value="53" name="skin"></td>
<td width="40" align="center"><input type="radio" value="54" name="skin"></td>
<td width="40" align="center"><input type="radio" value="55" name="skin"></td>
<td width="40" align="center"><input type="radio" value="56" name="skin"></td>
<td width="40" align="center"><input type="radio" value="57" name="skin"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="760" cellspacing="0" cellpadding="0">
<tr>
<td width="40"><img border="0" src="skin/58.gif"></td>
<td width="40"><img border="0" src="skin/59.gif"></td>
<td width="40"><img border="0" src="skin/60.gif"></td>
<td width="40"><img border="0" src="skin/61.gif"></td>
<td width="40"><img border="0" src="skin/62.gif"></td>
<td width="40"><img border="0" src="skin/63.gif"></td>
<td width="40"><img border="0" src="skin/64.gif"></td>
<td width="40"><img border="0" src="skin/65.gif"></td>
<td width="40"><img border="0" src="skin/66.gif"></td>
<td width="40"><img border="0" src="skin/67.gif"></td>
<td width="40"><img border="0" src="skin/68.gif"></td>
<td width="40"><img border="0" src="skin/69.gif"></td>
<td width="40"><img border="0" src="skin/70.gif"></td>
<td width="40"><img border="0" src="skin/71.gif"></td>
<td width="40"><img border="0" src="skin/72.gif"></td>
<td width="40"><img border="0" src="skin/73.gif"></td>
<td width="40"><img border="0" src="skin/74.gif"></td>
<td width="40"><img border="0" src="skin/75.gif"></td>
<td width="40"><img border="0" src="skin/76.gif"></td>
</tr>
<tr>
<td width="40" align="center"><input type="radio" value="58" name="skin"></td>
<td width="40" align="center"><input type="radio" value="59" name="skin"></td>
<td width="40" align="center"><input type="radio" value="60" name="skin"></td>
<td width="40" align="center"><input type="radio" value="61" name="skin"></td>
<td width="40" align="center"><input type="radio" value="62" name="skin"></td>
<td width="40" align="center"><input type="radio" value="63" name="skin"></td>
<td width="40" align="center"><input type="radio" value="64" name="skin"></td>
<td width="40" align="center"><input type="radio" value="65" name="skin"></td>
<td width="40" align="center"><input type="radio" value="66" name="skin"></td>
<td width="40" align="center"><input type="radio" value="67" name="skin"></td>
<td width="40" align="center"><input type="radio" value="68" name="skin"></td>
<td width="40" align="center"><input type="radio" value="69" name="skin"></td>
<td width="40" align="center"><input type="radio" value="70" name="skin"></td>
<td width="40" align="center"><input type="radio" value="71" name="skin"></td>
<td width="40" align="center"><input type="radio" value="72" name="skin"></td>
<td width="40" align="center"><input type="radio" value="73" name="skin"></td>
<td width="40" align="center"><input type="radio" value="74" name="skin"></td>
<td width="40" align="center"><input type="radio" value="75" name="skin"></td>
<td width="40" align="center"><input type="radio" value="76" name="skin"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="760" cellspacing="0" cellpadding="0">
<tr>
<td width="40"><img border="0" src="skin/77.gif"></td>
<td width="40"><img border="0" src="skin/78.gif"></td>
<td width="40"><img border="0" src="skin/79.gif"></td>
<td width="40"><img border="0" src="skin/80.gif"></td>
<td width="40"><img border="0" src="skin/81.gif"></td>
<td width="40"><img border="0" src="skin/82.gif"></td>
<td width="40"><img border="0" src="skin/83.gif"></td>
<td width="40"><img border="0" src="skin/84.gif"></td>
<td width="40"><img border="0" src="skin/85.gif"></td>
<td width="40"><img border="0" src="skin/86.gif"></td>
<td width="40"><img border="0" src="skin/87.gif"></td>
<td width="40"><img border="0" src="skin/88.gif"></td>
<td width="40"><img border="0" src="skin/89.gif"></td>
<td width="40"><img border="0" src="skin/90.gif"></td>
<td width="40"><img border="0" src="skin/91.gif"></td>
<td width="40"></td>
<td width="40"></td>
<td width="40"></td>
<td width="40"></td>
</tr>
<tr>
<td width="40" align="center"><input type="radio" value="77" name="skin"></td>
<td width="40" align="center"><input type="radio" value="78" name="skin"></td>
<td width="40" align="center"><input type="radio" value="79" name="skin"></td>
<td width="40" align="center"><input type="radio" value="80" name="skin"></td>
<td width="40" align="center"><input type="radio" value="81" name="skin"></td>
<td width="40" align="center"><input type="radio" value="82" name="skin"></td>
<td width="40" align="center"><input type="radio" value="83" name="skin"></td>
<td width="40" align="center"><input type="radio" value="84" name="skin"></td>
<td width="40" align="center"><input type="radio" value="85" name="skin"></td>
<td width="40" align="center"><input type="radio" value="86" name="skin"></td>
<td width="40" align="center"><input type="radio" value="87" name="skin"></td>
<td width="40" align="center"><input type="radio" value="88" name="skin"></td>
<td width="40" align="center"><input type="radio" value="89" name="skin"></td>
<td width="40" align="center"><input type="radio" value="90" name="skin"></td>
<td width="40" align="center"><input type="radio" value="91" name="skin"></td>
<td width="40" align="center"></td>
<td width="40" align="center"></td>
<td width="40" align="center"></td>
<td width="40" align="center"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="760" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" align="center"><input type="submit"
value="更改我的頭像" name="B1"
style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</center>
</div>
</form>

</body>

</html>