<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=trim(request.querystring("id"))
if not isnumeric(id) then 
	Response.Write "<script language=JavaScript>{alert('提示：["&id&"]輸入錯誤，ID請使用數字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名,保留2,mvalue from 用戶 where id=" & id & " and 介紹人='"&info(0)&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('這個人也不是你介紹來的，你小子想作什麼！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
toname=rs("姓名")
if rs("保留2")="扒皮"&Month(date()) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('喂,你想作什麼呀，你已經扒過皮了，還想扒啊?');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
yudian=rs("mvalue")
conn.Execute "update 用戶 set 保留2='"&"扒皮"&Month(date())&"' where id="&id
conn.Execute "update 用戶 set allvalue=allvalue+"&int(yudian*0.05)&" where id="&info(9)
rs.close
set rs=nothing
conn.close
set conn=nothing
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
sd(199)="<font color=FFD7F4>小道消息：["&info(0)&"]從("&toname&")身上扒到泡點:"&int(yudian*0.05)&"點，努力吧，多拉人！</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Response.Write "<script Language=Javascript>alert('提示：["&info(0)&"]從("&toname&")身上扒到泡點:"&int(yudian*0.05)&"點，努力吧，多拉人！');location.href = 'javascript:history.go(-1)';</script>"
%>
