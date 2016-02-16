<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=Trim(Request.QueryString("id"))
Session("ljjh_inthechat")<>"1"
Select Case id
Case "000"
nl="  對不起，程序所在目錄不是虛擬目錄，或設置有錯誤，Global.asa 沒有被執行。本程序需要虛擬目錄的支持！"
Case "001"
if session("nowinroom")<>"" then
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
ubb=UBound(onlinelist)
for i=1 to ubb step 6
if CStr(onlinelist(i+1))<>CStr(Session("ljjh_name")) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock

nl="本窗口即將被關閉。<br><br>造成此現像的原因主要有：<br>   由於網絡傳輸問題，導致你的機器與服務器在三分鐘內無法傳遞數據，系統將你作為超時而斷開了連接；<br>   你點擊了“換名登錄”重回登錄頁面，又沒有先關閉此窗口；<br>   你被踢出了聊天室。<br>   如果你使用本窗口進入不會出現這個問題，而用新窗口進入就出現這個問題的話，問題就出在你的機器：對彈出的新窗口無法繼承上級窗口的值。<br><br>解決方法：<br>  關閉此窗口，重新到登錄頁面輸入用戶名和密碼進行登錄。如果你是被踢出聊天室，則你方才所用的用戶名在 5 分鐘內不能使用。如果是第４種情況，可以試著清除ＩＥ的臨時文件，然後重新啟動機器。"
else
nl="本窗口即將被關閉。<br><br>造成此現像的原因主要有：<br>   由於網絡傳輸問題，導致你的機器與服務器在三分鐘內無法傳遞數據，系統將你作為超時而斷開了連接；<br>   你點擊了“換名登錄”重回登錄頁面，又沒有先關閉此窗口；<br>   你被踢出了聊天室。<br>   如果你使用本窗口進入不會出現這個問題，而用新窗口進入就出現這個問題的話，問題就出在你的機器：對彈出的新窗口無法繼承上級窗口的值。<br><br>解決方法：<br>  關閉此窗口，重新到登錄頁面輸入用戶名和密碼進行登錄。如果你是被踢出聊天室，則你方才所用的用戶名在 5 分鐘內不能使用。如果是第４種情況，可以試著清除ＩＥ的臨時文件，然後重新啟動機器。"
end if
Case "002"
nl="本窗口即將被關閉。<br><br>原因：<br>  你超過120分鐘沒有發言，為減輕服務器負擔，系統自動清除你所占用的資源。<br><br>解決方法：<br>  關閉此窗口，重新到登錄頁面輸入用戶名和密碼進行登錄，如果多次出現這種情況進不了那麼就點首頁的清除IE裡下載IE清除歷史記錄軟件運行就可以進江湖了！"
Case else
nl="  對不起，該出錯類型未被登記。"
End Select
Session.Abandon
nl="<p>" & nl & "</p>"%><html>
<head>
<title>出錯提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%>>
<script LANGUAGE="JavaScript">if(window!=window.top)top.location.href=location.href;</script>
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
<table border="0" bgcolor="#FF0099" cellspacing="0" cellpadding="2" width="350">
<tr>
<td width="342"><font color="FFFFFF"> 出錯提示</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="關閉"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value=" 確 定 " onclick=javascript:window.close()>
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