<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
grade=Int(info(2))
mypai=info(5)
inthechat=Session("ljjh_inthechat")
if grade<10 then Response.Redirect "../error.asp?id=482"
if inthechat<>"1" then Response.Redirect "manerr.asp?id=482"
bombname=Trim(Request.QueryString("id"))
if bombname="" then Response.Redirect "../error.asp?id=481"
if CStr(bombname)=CStr(info(0)) then Response.Redirect "../error.asp?id=481"%><html>
<head>
<title>炸彈操作</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="0099FF"><font size="6" face="隸書" color="#000000">炸 彈 操 作</font></font></h1>
</div>
<hr noshade size="1" color=009900 width="70%">
<div align="center">
<table width="380" border="0" cellspacing="2" cellpadding="2">
<tr>
<td><b> 注意事項 </b> <br>
必須輸入扔炸彈的原因才能投放炸彈，被炸的對像的機器將不停地彈出新窗口，直到系統資源耗盡、死機。<br>
炸彈操作會被記錄在“聊務公開”欄中，供廣大聊友監督，因此，請勿隨意炸人！ </td>
</tr>
</table>

</div>
<hr noshade size="1" color=009900 width="70%">
<br>
<table border="1" cellspacing="0" cellpadding="3" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" width="380">
<form method="post" action="manbombok.asp">
<tr>
<td>
<table border="0" cellpadding="4" width="100%" cellspacing="0">
<tr>
<td align="right" width="26%">轟炸用戶名：</td>
<td width="74%"><font color="#FF0000"><%=bombname%>
<input type="hidden" name="bombname" value="<%=bombname%>">
</font></td>
</tr>
<tr>
<td align="right" width="26%">轟炸的原因：</td>
<td width="74%">(&lt;=60字符) </td>
</tr>
<tr>
<td align="right" width="26%">
<select name="select" onchange="javascript:document.forms[0].bombwhy.value=this.value;document.forms[0].bombwhy.focus();">
<option value="" selected>自填</option>
<option value="不受歡迎人。">不歡迎</option>
<option value="所取的名字十分不雅。">不雅ID</option>
<option value="亂刷屏，警告又不聽。">亂刷屏</option>
<option value="在聊天室散布有悖倫理道德的言論。">不道德</option>
<option value="不遵守聊天規則，進行人身攻擊。">罵人</option>
<option value="在聊天室散布違反國家法律法規的言論。">違法</option>
</select>
：</td>
<td width="74%">
<input type="text" name="bombwhy" maxlength="60" size="40">
</td>
</tr><%if info(0)="江湖總站" then%>
<tr>
<td align="right" width="26%">管理員選項：</td>
<td width="74%">
<input type="radio" name="logok" value="1" checked>記入聊務公開欄
<input type="radio" name="logok" value="0">不記入聊務公開欄</td>
</tr><%end if%>
</table>
<div align="center">
<input type="submit" name="bombok" value="轟炸">
<input type="button" value="放棄" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<br>
<hr noshade size="1" color=009900 width="70%">
<div align=center class=cp></div>
</body>
</html> 