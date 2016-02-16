<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../images/8.jpg">
<%my=info(0)
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 and money<>10000000 then 
	Response.Write "<script Language=Javascript>alert('你想作什麼？！');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩 from 用戶 where id="& info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	coon.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
if rs("銀兩")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你沒有錢呀，你想作什麼？！');window.close();</script>"
	response.end
end if
conn.execute "update 用戶 set 道德=道德+"&int(money/500)&",銀兩=銀兩-"& money &",操作時間=now() where id="&info(9)
rs.close
rs.open "select 等級,道德加,道德 from 用戶 where id="&info(9),conn
ddj=(rs("等級")*1440+2200)+rs("道德加")
if rs("道德")>ddj then
conn.execute "update 用戶 set 道德=" & ddj & "  where id="&info(9)
end if
Response.Write "<script Language=Javascript>alert('您向你的孤兒院捐獻了"& money &"兩，道德上漲"&int(money/500)&"點！');window.close();</script>"
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
sd(199)="<font color=FFD7F4>【愛心奉獻】["&info(0)&"]</font>大俠為靈劍江湖的孤兒院的孤兒們捐獻了"& money &"兩，積點德啦^_^" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
