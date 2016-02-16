<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(2)<10  then Response.Redirect "../error.asp?id=439"
id=request("id")
rs.open ("SELECT 被告,原告,要求,標題 FROM 申冤 WHERE ID=" & id),conn,0,1
bg=rs("被告")
yg=rs("原告")
result=rs("要求")
mess=rs("標題")

if result="罰款一萬" then
sql="update 用戶 set 銀兩=銀兩-10000 where 姓名='"&bg&"'"
conn.Execute(sql)
sql="update 用戶 set 銀兩=銀兩+10000 where 姓名='"&yg&"'"
conn.Execute(sql)
end if
if result="罰款十萬" then
sql="update 用戶 set 銀兩=銀兩-100000 where 姓名='"&bg&"'"
conn.Execute(sql)
sql="update 用戶 set 銀兩=銀兩+100000 where 姓名='"&yg&"'"
conn.Execute(sql)
end if
if result="坐牢" then
sql="update 用戶 set 狀態='牢',登錄=now()+1/144 where 姓名='" & bg & "'"
conn.execute sql
end if

sql="update 申冤 set 結果='1' where id=" & id & ""
conn.execute sql
newer="江湖人士<font color=FFD7F4>" & yg & "</font>狀告<font color=FFD7F4>" & bg & "</font>標題：<font color=red>{"&mess&"}</font>,要求<font color=red>" & result & "</font>,因證據充分，通過申議！"
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
sd(199)="<font color=FFD7F4>【申冤判決】</font>"& newer &""
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

'踢人
sub boot(bg)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(bg) then
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
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
conn.close
set rs=nothing	
set conn=nothing
Response.Redirect "manage.asp"
%>