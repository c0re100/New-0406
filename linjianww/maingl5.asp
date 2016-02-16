<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM 用戶 WHERE 門派<>'官府' order by mvalue desc"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('目前沒有月積分排行！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
else
do while not rs.eof
conn.execute "update 用戶 set 法力=法力+50 where 姓名='"&rs("姓名")&"'"
rs.movenext
file=file+1
if file>20 then Exit Do
loop
end if
set rs=nothing
conn.close
set conn=nothing
mess="月積分排行前20名獎勵法力50元，大家祝賀，繼續努力呀！"
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
sd(198)="對"
sd(199)="<font size=3 color=red>【獎勵消息】</font><font color=FFD7F4 size=3>"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('現在已將所有月積分前20名的玩家獎勵了50點法力！！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>