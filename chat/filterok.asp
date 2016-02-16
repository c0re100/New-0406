<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if nickname="" or Session("ljjh_inthechat")<>"1" then Response.Redirect "close.asp"
filname=Request.Form("filtername")
filname=" "&filname&","
clearall=Request.Form("clearall")
if clearall="取消" then filname=" ,"
Session("ljjh_filname")=filname
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>屏蔽討厭鬼</title>
<style type='text/css'>
body {font-size:12pt;}
td {font-size:10.5pt;line-height:120%}
input{font-size:9pt}
select{font-size:9pt}
textarea{font-size:9pt}
.p9{font-size:9pt}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" topmargin="0">
<table width="154" border="1" cellspacing="0" cellpadding="4" bgcolor="#EEFFEE" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="267">
  <form>
<tr align="center">
<td>
<font color="red">屏蔽結果</font><br>
<%if filname=" ," then%><p>關閉此功能</p><%else%>
<font class="p9">下列用戶的發言將被屏蔽，不會再出現在你的屏幕上。</font>
<p><font color=blue><%=filname%></font></p><%end if%>
<input type="button" value="返回" onclick="javascript:history.go(-1)">
</td>
</tr>
</form>


</table>
</body>
</html> 