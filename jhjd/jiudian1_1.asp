<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=request("id")
%><!--#include file="dadata.asp"-->
<%
rs.open "select * from 用戶 where id="&info(9),conn
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能訂購酒宴"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "SELECT * FROM 宴會 where ID=" & id,conn
wu=rs("宴會名")
yin=rs("售價")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
jb=rs("金幣")
sl=rs("數量")%>
<%
rs.close
rs.open "select * from 用戶 where id="&info(9),conn
if yin>=rs("銀兩") then
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
	conn.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('不能定酒宴，原因：你的銀兩不夠了');history.back();</script>"
	response.end
end if
	conn.execute "update 用戶 set 銀兩=銀兩-" & yin & " where 姓名='" & info(0) & "'"
	rs.close
	rs.open "select * from 宴會列表 where 宴會名='" & wu & "' and 擁有者='" & info(0) & "'",conn
if rs.eof or rs.bof then
	connt.execute "insert into 宴會列表(宴會名,擁有者,類型,內力,體力,金幣,數量,售價,時間) values ('"&wu&"','"&info(0)&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
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
	sd(199)="<font color=FFFDAF>【好消息】"&info(0)&"在江湖大酒店的"&wu&"廳舉行<font color=red>※"&lx&"※</font>宴會，大家都快去呀，晚了就沒的喫了。。。</font>"
	sd(200)=session("nowinroom")
	Application("ljjh_sd")=sd
	Application.UnLock
	Response.Redirect "jd.asp"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('不能定酒宴，原因：你已定了同樣規格的酒席！');history.back();</script>"
	response.end
end if
%>
