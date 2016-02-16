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
if info(2)<10 then Response.Redirect "../error.asp?id=439"
result=request("result")
id=request("id")
sql="update 申冤 set 結果='0' where id=" & id & ""
conn.execute sql
rs.open ("SELECT 被告,原告,要求,標題 FROM 申冤 WHERE ID=" & id),conn,0,1
bg=rs("被告")
yg=rs("原告")
result=rs("要求")
mess=rs("標題")
newer="江湖人士<font color=FFD7F4>" & yg & "</font>狀告<font color=FFD7F4>" & bg & "</font>標題：<font color=red>{"&mess&"}</font>,要求<font color=red>" & result & "</font>,因證據不足、理由不充分，未通過！"
rs.close
conn.close
set rs=nothing	
set conn=nothing
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
sd(199)="<font color=FFD7F4>【申冤判決】</font>"& newer 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

Response.Redirect "manage.asp"
%>