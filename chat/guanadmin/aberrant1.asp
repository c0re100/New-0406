<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<8  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim conn,rs,rsuser
on error resume next
opt=Request.QueryString("opt")
st=Request.QueryString("touser")
set rs=server.createobject("adodb.recordset")
rs.open ("select * from 用戶 where 姓名='"&st&"'"),conn,0,1
if rs.bof or rs.eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Redirect "../error.asp?id=446"
else
grade1=rs("grade")
stt=rs("id")
if info(2)<=grade1 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
alert "對不起，對方是是高級管理人員，你的管理等級比對方低！"
window.close()
</script>
<%
response.end
end if
if st="江湖總站" and st="呆呆"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
alert "對不起，對方是特赦人員！"
window.close()
</script>
<%
response.end
end if
 if opt="逮捕" then
call boot(st)
conn.execute("update 用戶 set 狀態='獄',登錄=now() where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','逮捕',now(),'特別管理行動')"
msg="<font color=FF0000>【特別管理行動】</font>天空中突然伸出一隻巨手，把不聽話的小雞<font color=FFC01F>"&st&"</font>抓入了官府大牢。(逮捕)[執行人員="&info(0)&"]"
end if
if opt="入獄" then
conn.execute("update 用戶 set 狀態='牢',登錄=now()+1/84 where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','坐牢',now(),'特別管理行動')"
msg="<font color=FF0000>【特別管理行動】</font>隨著一聲威武的吆喝，官府的人如狼似虎的將不聽話的小雞<font color=FFC01F>"&st&"</font>關進了重犯大牢。(坐牢)[執行人員="&info(0)&"]"

call boot(st)
end if
if opt="踢人" then
'conn.execute("update 用戶 set 狀態='牢',登錄=now()+1/84 where id="&stt)
'conn.execute("update 用戶 set 狀態='正常',登錄=now() where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','踢人',now(),'特別管理行動')"
msg="<font color=FF0000>【踢人】</font>隻聽砰的一聲，<font color=FFC01F>"&st&"</font>被踢出聊天室。[執行人員="&info(0)&"]"
call boot(st)
end if
if opt="有期徒刑" then
conn.execute("update 用戶 set 狀態='牢',登錄=now()+1/4 where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','坐牢',now(),'特別管理行動')"
msg="<font color=FF0000>【特別管理行動】</font><font color=FFC01F>"&st&"</font>惡事做多被判入牢24小時。(有期徒刑)[執行人員="&info(0)&"]"
call boot(st)
end if
if opt="監禁" then
conn.execute("update 用戶 set 狀態='監禁',登錄=now()+8000 where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','監禁',now(),'特別管理行動')"
msg="<font color=FF0000>【特別管理行動】</font><font color=FFC01F>"&st&"</font>惡事做多被帶入了十八層地獄。(無期徒刑)[執行人員="&info(0)&"]"
call boot(st)
end if
if opt="睡覺" then
conn.execute("update 用戶 set 狀態='眠',登錄=now()+8000 where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','監禁',now(),'特別管理行動(睡覺2024年')"
msg="<font color=FF0000>【長眠】</font><font color=FFC01F>"&st&"</font>惡事做多被帶入了十八層地獄，囚禁到2024年。[執行人員="&info(0)&"]"
call boot(st)
end if
if opt="炸彈" then
conn.execute("update 用戶 set 狀態='牢',登錄=now()+8888 where id="&stt)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & st & "','監禁',now(),'特別管理行動(人間蒸發')"
msg="<font color=FF0000>【人間蒸發】</font>一陣電閃雷鳴，<font color=FFC01F>"&st&"</font>因違反江湖規矩，被炸出聊天室....[執行人員="&info(0)&"]"
'conn.execute("delete * from 用戶 where id="&stt)
call zhadan(st)
end if
if opt="寶物" then
conn.execute("update 用戶 set 保護=false,寶物修練=0,寶物='靈劍水晶石' where id="&stt)
msg="江湖消息：<a href=javascript:parent.sw('[" & st & "]'); target=f2><font color="& addwordcolor & "><b>" & st & "</b></font></a>現在擁有江湖至寶<font color=red>靈劍水晶石</font>,這寶物是江湖總站尋便大江南北，才辛辛苦苦找到的啊，唉。。。各位俠士互相爭奪！[執行人員="&info(0)&"]"
end if
if opt="獎勵" then
conn.execute("update 用戶 set 存款=存款+30000000 where id="&stt)
'msg="<font color=FFC01F>【站長獎勵】</font><font color=FFC01F>"&st&"</font>表現優秀，舉報有功,站長親臨<font color=FFC01F>"&st&"</font>的家中送上江湖貢獻獎勵金！"
end if
if opt="罰款" then
conn.execute("update 用戶 set 存款=存款-30000000 where id="&stt)
msg="<font color=FFC01F>【站長罰款】</font><font color=FFC01F>"&st&"</font>表現惡劣，嚴重搗亂江湖，影響極壞決定罰款3000萬以示懲戒！"
end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

'踢人
sub boot(st)
Application.Lock
onlinelist=Application("ljjh_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(st) then
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
Application("ljjh_onlinelist")=listnull
else
Application("ljjh_onlinelist")=newonlinelist
end if
Application("ljjh_useronlinename")=useronlinename
Application("ljjh_chatrs")=onliners
Application.UnLock
end sub
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
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="對"
sd(199)="<font color=FFC01F>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.Lock
'炸彈
sub zhadan(st)
Application.Lock
onlinelist=Application("ljjh_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(st) then
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
Application("ljjh_onlinelist")=listnull
else
Application("ljjh_onlinelist")=newonlinelist
end if
Application("ljjh_useronlinename")=useronlinename
Application("ljjh_chatrs")=onliners
Application.UnLock
end sub
Response.redirect "onlinelist1.asp"
%>