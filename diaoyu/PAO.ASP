<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
conn.execute "update 用戶 set 操作時間=now() where id="&info(9)

conn.close
set conn=nothing
Session("diaoyu")=true
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=#ff0000>消息</font>"& info(0) &"在河邊釣魚：由於太不專心，給魚跑了。。。"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>魚跑了</title></head>
<body background="../jhimg/bk_hc3w.gif">
<div align="center">
<p>&nbsp;</p>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="200">
<tr>
<td bgcolor=#C6BD9B>
<table>
<tr>
<td valign="top">
<div align="center">
很可惜，你的魚溜走了......下次再來釣過吧。
</div>
</table>
<div align="center"><br>
<input type=button value=關閉窗口 onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>