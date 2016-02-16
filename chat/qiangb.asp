<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
if Application("ljjh_qiang")=0 then
	Response.Write "<script Language=Javascript>alert('提示：你來晚了，東西已經被別人搶走了...');</script>"
	response.end
end if
Application.Lock
Application("ljjh_qiang")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 物品名,銀兩,類型,內力,體力,等級,堅固度,說明,sm,會員 FROM 物品 where id=" & ID,conn
wu=rs("物品名")
yin=rs("銀兩")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
say=rs("說明")
sm=rs("sm")
dj=rs("等級")
jg=rs("堅固度")
rs.close
'物品
rs.open "select id from 物品 where 物品名='"&wu&"' and 擁有者="& info(9),conn
If Rs.Bof OR Rs.Eof then
	sql="insert into 物品(物品名,擁有者,類型,內力,體力,堅固度,數量,銀兩,等級,說明,sm,會員) values ('"&wu&"',"&info(9)&",'"&lx&"',"&nl&","&tl&","&jg&",1,"&yin&","&dj&",'"&say&"',"&sm&","&aaao&")"
	conn.execute sql
else
	id1=rs("id")
	sql="update 物品 set 數量=數量+1,會員="&aaao&" where id="&id1
	conn.execute sql
end if
conn.execute "update 物品 set 數量=數量-1,會員="&aaao&" where id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你搶到了一樣東西，快看屏幕！');</script>"
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
sd(195)=info(0)
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="對"
sd(199)="<font color=9FDF70>【搶尸】"&info(0)&"拖展飛龍探雲手搶到一個["&wu&"]。</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%> 