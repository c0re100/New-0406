<%@ LANGUAGE=VBScript%>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 物品 where 數量>0 and 物品名='免罪卡' and 擁有者="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('你免罪卡嗎？！');location.href = 'listlao.asp';}</script>"
	Response.End
end if
conn.Execute "update 物品 set 數量=數量-1 where 擁有者="&info(9)&" and 物品名='免罪卡'"
conn.execute "update 用戶 set 狀態='正常',登錄=now() where 姓名='" & name & "'"
msg="<font color=green>【免罪卡】</font><font color=ffffff>"&info(0)&"</font>偷偷使用了免罪卡把在牢裡的<font color=blue>" & name & "</font>給保釋了出來！"
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
sd(196)="660099"
sd(197)="660099"
sd(198)="對"
sd(199)="<font color=red>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "listlao.asp"
%>