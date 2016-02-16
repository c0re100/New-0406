<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
Select Case id
Case "100"
nl="成功修改密碼，新密碼為 <font color=red>" & Request.QueryString("new") & "</font>，請記牢。"
Case "101"
nl="自殺操作完成！用戶名 <font color=red>" & Request.QueryString("name") & "</font> 已經被刪除，該用戶名的所有記錄及權限均已消失。"
Case "110"
nl="個人信息已經修改成功。"
Case "111"
nl="你改名操作完成！"
Case "200"
nl="消息已經成功發送給<font color=red>" & Request.QueryString("name") & "</font>。可以在消息列表中查看該消息是否已經被取出。"
Case "210"
nl="動作已經添加完成。"
Case "300"
nl="數據庫備份完成，請到 backup 目錄下載 global.mdb 進行壓縮。"
Case "301"
nl="已經用壓縮後的數據庫覆蓋舊數據庫，確認證確後，請刪除該備份數據庫，以防被他人下載。"
Case "302"
nl="備份數據庫刪除完成。"
Case "600"
nl="結婚登記成功！收取費用1000兩！"
Case "601"
nl="經過漫長的洗浴，你發現自己清爽了不少，精神也好多了，也許該到家裡睡一覺會更加開心吧。"
Case "700"
nl="恭喜你！數據庫更新成功！"
Case "701"
nl="門派資料已成功修改！"
Case "702"
nl="專用包設置完成！請刪除 setup.asp 文件，以免重復運行！"
Case "703"
nl="掌門被開除了！"
Case "704"
nl="刪除成功！！"
Case "705"
nl="江湖告示：官符救災工作順利完工！"
Case "706"
nl="江湖告示：官符收繳百姓藥器的工作順利完工！"
Case "707"
nl=""
Case else
nl="對不起，該類型未被登記。"
End Select
nl="<p>  " & nl & "</p>"%><html>
<head>
<title>操作成功</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false  background=JHIMG/Bk_hc3w.gif leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#dcd8d0">
<tr>
<td>
<table border="0" cellspacing="0" cellpadding="2" width="370" background="jhimg/title.jpg">
<tr>
<td width="344"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">操作成功</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="22">
<tr>
<td width="18"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="返回"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="348" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="267"> <%=nl%> </td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<%if id="200" then
Response.Write "<input type='button' name='ok' value=' 查看列表 ' onclick=javascript:top.location.href='webicqlist.asp'>"
else
Response.Write "<input type='button' name='ok' value=' 返 回 ' onclick='javascript:history.go(-1)'>"
end if%>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html> 