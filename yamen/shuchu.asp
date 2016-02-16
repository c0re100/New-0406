<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
my=info(0)
id=request("id")
'校驗用戶
rs.open "select grade,ID from 用戶 where id="&info(9),conn
If Rs.Bof OR Rs.Eof Then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('你不是江湖中人！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if  rs("grade")<6 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('你無此權力！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
UID=rs("ID")
rs.close
rs.open "select id from 用戶 where 狀態='牢' or 狀態='獄'  and id=" & id
if  rs.eof and  rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('沒這個人！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
conn.execute "update 用戶 set 狀態='正常',登錄=now() where id=" & id
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="【釋放】"
'sd(199)="<font color=red>【釋放】" & rs("姓名") & "被無罪釋放(執行官["&info(0)&"])</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('釋放成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
