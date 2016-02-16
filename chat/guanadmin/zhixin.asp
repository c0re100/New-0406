<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(2)<8  then Response.Redirect "../../error.asp?id=425"
pass1=Request.Form("pass1")
pass11=Request.Form("pass11")
ljgl=Request.Form("ljgl")
if InStr(pass1,"'")<>0 or InStr(pass1,"`")<>0 or InStr(pass1,"=")<>0 or InStr(pass1,"-")<>0 or InStr(pass1,",")<>0 or InStr(pass1,"\")<>0 or InStr(pass1,"/")<>0 or InStr(pass1,"|")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');history.back();}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT grade FROM 用戶 where 姓名='"&pass1&"' and 狀態='正常'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('抱歉你所要抓的人在數據庫裡沒有！請查看名字是否正確！');history.back();</script>"
Response.End
else
 if rs("grade")>info(2) then
	
 Response.Write "<script language=javascript>alert('抱歉你所要抓的人管理級比你高，你抓不了的！');history.back();</script>"
Response.End
 end if 
select case ljgl
case "逮捕"
sql="update 用戶 set 狀態='獄',登錄=now() where 姓名='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" &pass1& "','逮捕',now(),'" &pass11& "')"
msg="<font color=FF0000>【特別管理行動】</font>天空中突然伸出一隻巨手，把不聽話的小雞<font color=FFFDAF>"&pass1&"</font>抓入了官府大牢。(逮捕)（原因：<font color=FFFDAF>"&pass11&"</font>）[執行人員="&info(0)&"]"
call boot(pass1)
case "坐牢"
sql="update 用戶 set 狀態='牢',登錄=now()+1/84 where 姓名='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" &pass1& "','坐牢',now(),'" &pass11& "')"
msg="<font color=FF0000>【特別管理行動】</font>隨著一聲威武的吆喝，官府的人如狼似虎的將不聽話的小雞<font color=FFFDAF>"&pass1&"</font>關進了重犯大牢。(坐牢)（原因：<font color=FFFDAF>"&pass11&"</font>）[執行人員="&info(0)&"]"
call boot(pass1)
case "監禁"
sql="update 用戶 set 狀態='監禁',登錄=now()+8000 where 姓名='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" &pass1& "','監禁',now(),'" &pass11& "')"
msg="<font color=FF0000>【特別管理行動】</font><font color=FFFDAF>"&pass1&"</font>惡事做多被帶入了十八層地獄。(監禁)（原因：<font color=FFFDAF>"&pass11&"</font>）[執行人員="&info(0)&"]"
call boot(pass1)
'case "睡覺"
'sql="update 用戶 set 狀態='眠',登錄=now()+8000  where 姓名='"&pass1&"'"
'Set Rs=conn.Execute(sql)
'conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" &pass1& "','監禁',now(),'" &pass11& "')"
'msg="<font color=FF0000>【特別管理行動】</font><font color=FFFDAF>"&st&"</font>惡事做多被帶入了十八層地獄，囚禁到2024年。[執行人員="&info(0)&"]"
'call boot(pass1)
case "寶物" 
sql="update 用戶 set 保護=false,寶物修練=0,寶物='靈劍水晶石' where 姓名='"&pass1&"'"
Set Rs=conn.Execute(sql)
msg="江湖消息：<a href=javascript:parent.sw('[" & pass1 & "]');target=f2><font color=FFFDAF><b>" & pass1 & "</b></font></a>現在擁有江湖至寶<font color=FFFDAF>靈劍水晶石</font>,這寶物是江湖總站尋便大江南北，才辛辛苦苦找到的啊，唉。。。各位俠士互相爭奪！"
'conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" &pass1& "','放寶物',now(),'" &pass11& "')"
case "炸彈"
sql="update 用戶 set 狀態='牢',登錄=now()+8888 where 姓名='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" &pass1& "','監禁',now(),'炸彈:" &pass11& "')"
msg="<font color=FF0000>【人間蒸發】</font>一陣電閃雷鳴，<font color=FFFDAF>"&pass1&"</font>因違反江湖規矩，被炸出聊天室....（原因：<font color=FFFDAF>"&pass11&"</font>）[執行人員="&info(0)&"]"
Application.Lock
Application("ljjh_bombname")=Application("ljjh_bombname") & " " & pass1 & " "
Application.UnLock
call zhadan(pass1)
end select

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
sd(199)="<font color=FFFDAF>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.Lock
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script language=javascript>alert('["&ljgl&"]：["&pass1&"]執行操作成功！');history.back();</script>"
end if
'踢人
sub boot(pass1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(pass1) then
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
end sub
'炸彈
sub zhadan(pass1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(pass1) then
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
end sub
%>