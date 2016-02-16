<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
dim conn,rs
%>

<html>
<head>
<title>提交狀紙</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../css.css">
<script LANGUAGE="javascript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.topic.value=="")
{
alert("狀紙題目沒有填！")
document.FrmAddLink.topic.focus()
return false
}
else if(document.FrmAddLink.name.value=="")
{
alert("狀告何人沒有填！")
document.FrmAddLink.name.focus()
return false
}
else if(document.FrmAddLink.play.value=="")
{
alert("對被告的要求沒有填！")
document.FrmAddLink.play.focus()
return false
}
else if(document.FrmAddLink.text.value=="")
{
alert("狀詞沒有填！")
document.FrmAddLink.text.focus()
return false
}
}

//-->
</script>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="569" valign="bottom"><img src="../images/sy5.gif" width="58" height="23"></td>
<td width="209"><img src="../images/sy1.jpg" width="227" height="83"></td>
</tr>
<tr>
<td background="../images/sy4.gif" width="569" align="right"><img src="../images/sy3.gif" width="94" height="42"></td>
<td background="../images/sy4.gif" width="209"><img src="../images/sy2.jpg" width="227" height="42"></td>
</tr>
</table>
<form action=newplan.asp method=post  LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
&nbsp;&nbsp;&nbsp;<b><font color="#FF6633">提交狀紙</font></b>
<table border=0 cellspacing=0 cellpadding=0 align=center width="590">
<tr>
<td height="26" colspan="2">狀紙題目:
<input name=topic size=50 maxlength="30">
最長30字</td>
</tr>
<tr>
<td height="17" colspan="2">狀告何人:
<input name=name size=10 maxlength="10">
請寫完整的姓名，否則不受理</td>
</tr>
<tr>
<td colspan="2" height="14">
<div align="left">如何處理:
<select name="play" size="1">
<option value="罰款一萬" selected>罰款一萬</option>
<option value="罰款十萬">罰款十萬</option>
<option value="坐牢">坐 牢</option>
</select>
罰款全數歸受害人所有</div>
</td>
</tr>
<tr align="center">
<td colspan="2">
<div align="center"><b><img src="../jhimg/bbs/edit.gif" width="18" height="13">
下面寫你的狀詞：<br>
</b>
<textarea name=text cols=50 rows=10></textarea>
</div>

<input type=submit value=發表 name="submit" style="background-color:FF0000;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">
&nbsp;&nbsp;&nbsp;&nbsp;
<input type=reset value=清除 name="reset" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">
</td>
</tr>
</table>
</form>
</body>
</html>
