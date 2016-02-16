<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
chatroombgimage=Application("hxf_c_chatroombgimage")
chatroombgcolor=Application("hxf_c_chatroombgcolor")
nl=""
Select Case id
Case "000"
 nl="程序所在目錄不是虛擬目錄，或設置有錯誤，Global.asa 沒有被執行。本程序需要虛擬目錄的支持！"
Case "100"
 nl="對不起，你尚未登錄或已經超時斷開連接，請回到登錄頁面重新輸入用戶名和密碼進行登錄。"
Case "101"
 nl="必須擁有 3 級以上權限才能添加自創的動作。"
Case "102"
 nl="各個項目均不能放空白，請填寫完整！"
Case "103"
 nl="動作內容必須是以“//”開頭的語句！"
Case "104"
 nl="動作內容不能出現“//##”，這將導致發言後姓名出現重復，請將 ## 去掉！"
Case "105"
 nl="動作名稱或動作內容超過限制的長度！"
Case "106"
 nl="動作內容含有“%%”，動作類型應改為“1 對別人”！"
Case "107"
 nl="動作類型為“1 對別人”，動作內容中卻沒有出現“%%”！"
Case "108"
 nl="添加的動作內容中，不能包含半角的“\”、“∣”、“" & chr(34) & "”！"
Case "109"
 nl="該內容已經存在，不必重復添加。"
Case "120"
 nl="你沒有清理“聊務公開”的權限。"
Case "121"
 nl="沒有超過 7 天的記錄，不能清除。"
Case "130"
 nl="你沒有“帳號管理”的權限。"
Case "131"
 nl="沒有此類帳號可供刪除。"
Case "132"
 nl="請輸入用戶名。"
Case "133"
 nl="在已“自殺”的帳號中沒有找到該用戶名，不能恢復。"
Case "134"
 nl="不能恢復用戶名：<font color=red>" & Request.QueryString("name") & "</font>，因為系統中已有相同的用戶名存在。如果你確實想要恢復該用戶名，則請用“刪除帳號”功能，先刪除系統中相同的用戶名後再恢復該用戶名。"
Case "135"
 nl="請輸入欲刪除的用戶名。"
Case "136"
 nl="用戶名不存在，不能刪除。"
Case "137"
 nl="用戶名不存在。"
Case "138"
 nl="該用戶名本已永久保留，不必重復操作。"
Case "139"
 nl="該用戶名未被永久保留，不能取消。"
Case "140"
 nl="舊用戶名和新用戶名均不能為空。"
Case "141"
 nl="不能將用戶名改為：<font color=red>" & Request.QueryString("name") & "</font>，因為系統中已有相同的用戶名存在。如果你確實想要改成這個用戶名，則請用“刪除帳號”功能，先刪除系統中相同的用戶名。"
Case "142"
 nl="用戶名和新密碼均不能為空。"
Case "150"
 nl="你沒有“數據壓縮”的權限。"
Case "151"
 nl="數據庫尚未關閉，不必重新打開數據庫。"
Case "152"
 nl="數據庫尚未打開，不必關閉數據庫。"
Case "160"
 nl="請輸入要搜索的用戶名。"
Case "200"
 nl="你沒有“踢人”的權限。"
Case "201"
 nl="你不在聊天室中，不能執行“踢人”操作。"
Case "202"
 nl="請指定要踢出的對像。"
Case "203"
 nl="不得無故踢人，請輸入原因。"
Case "204"
 nl="用戶名：<font color=red>" & Request.QueryString("kickname") & "</font> 不在聊天室中，不必再踢了。"
Case "205"
 nl="呵呵，毛病，踢自己玩啊。"
Case "210"
 nl="你沒有“IP管理”的權限。"
Case "211"
 nl="你不在聊天室中，不能進行“IP管理”。"
Case "212"
 nl="請指定要封鎖的 IP。"
Case "213"
 nl="封鎖自己的 IP？別傻了！"
Case "214"
 nl="不得無故封鎖 IP，請輸入原因。"
Case "215"
 nl="該 IP 已經被封鎖了，不能重復封鎖。"
Case "216"
 nl="請輸入解鎖該 IP 的原因。"
Case "217"
 nl="該 IP 未被封鎖，不能進行解鎖。"
Case "218"
 nl="請指定要解鎖的 IP。"
Case "219"
 nl="要封鎖的IP與用戶名不對應。"
Case "220"
 nl="你沒有“炸彈操作”的權限。"
Case "221"
 nl="你不在聊天室中，不能進行“炸彈操作”。"
Case "222"
 nl="請指定要轟炸的對像。"
Case "223"
 nl="啊，你不是廁所裡打燈籠––找死吧。"
Case "224"
 nl="不得無故亂扔炸彈，請輸入原因。"
Case "225"
 nl="用戶名：<font color=red>" & Request.QueryString("bombname") & "</font> 不在聊天室中，炸不到了。"
Case "230"
 nl="你沒有更改“系統參數”的權限。"
Case "231"
 nl="新值與舊值完全相同，不必進行更改。"
Case "232"
 nl="輸入的值不合法。"
Case "240"
 nl="你沒有進行“級別管理”的權限。"
Case "241"
 nl="你不在聊天室中，不能進行“級別管理”。"
Case "242"
 nl="你沒有執行“升級操作”的權限。"
Case "243"
 nl="你沒有執行“降級操作”的權限。"
Case "244"
 nl="用戶名不能為空。"
Case "245"
 nl="找不到用戶名：<font color=red>" & Request.QueryString("username") & "</font>。"
Case "246"
 nl="該用戶名的級別不比你低，不能對其進行操作。"
Case "247"
 nl="選定的級別值不合法。"
Case "248"
 nl="請輸入原因。"
Case "250"
 nl="你沒有查看“HTML妙用”的權限。"
Case "255"
 nl="你沒有更改“站長公告”的權限。"
Case "256"
 nl="內容不能為空。"
Case "260"
 nl="你沒有“動作管理”的權限。"
Case "261"
 nl="找不到該動作。"
Case "270"
 nl="你沒有“留言管理”的權限。"
Case "271"
 nl="該留言不存在。"
Case "280"
 nl="你沒有“調整級別”的權限。"
Case "281"
 nl="你不在聊天室中不能“調整級別”。"
Case "282"
 nl="用戶名不能為空。"
Case "283"
 nl="請輸入調整級別的原因。"
Case "300"
 nl="你沒有管理“投票系統”的權限。"
Case "301"
 nl="不符合投票條件，不能投票。"
Case "302"
 nl="格式錯誤。"
Case "303"
 nl="請選擇你支持的候選人。"
Case "304"
 nl="候選人不存在。"
Case "305"
 nl="候選人不能為空。"
Case "306"
 nl="候選人已經存在，不能重復添加。"
Case "310"
 nl="你沒有“永久封鎖”ＩＰ的權限。"
Case "311"
 nl="格式錯誤。"
Case "320"
 nl="IP不能為空。"
Case else
 nl="對不起，該出錯類型未被登記。"
End Select
nl="<p>.2' g Sl[dBu<A\>'%><html>
<head>
<title>出錯提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%> leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center"> 
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
              <table border="0" bgcolor="#3399FF" cellspacing="0" cellpadding="2" width="350">
                <tr>
<td width="342"><font color="FFFFFF"> 出錯提示</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:history.go(-1)" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="關閉"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr> 
                  <td width="59" align="center" valign="top"><font face="Wingdings" color="#0066FF" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value=" 確 定 " onclick=javascript:history.go(-1)>
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
 