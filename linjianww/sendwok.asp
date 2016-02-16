<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(0)<>"江湖總站" then Response.Redirect "../error.asp?id=439"
a=request.form("a") '對方ID
b=request.form("b") '物品名
c=request.form("c") '物品數量
d=request.form("d") '類型
e=trim(request.form("e")) '說明
f=request.form("f") '內力
g=request.form("g") '體力
h=request.form("h") '攻擊,
i=request.form("i") '防禦
j=request.form("j") '銀兩
k=request.form("k") '堅固度
m=request.form("m") '等級
n=request.form("n") '會員
l=request.form("l") 'sm
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

rs.open "SELECT id FROM 用戶 where 姓名='"&a&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，沒有該用戶！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
idd=rs("id")
rs.close
rs.open "select id from 物品 where 物品名='"&b&"' and 擁有者="&idd,conn
if rs.eof and rs.bof then
conn.execute "insert into 物品(物品名,擁有者,類型,說明,內力,體力,攻擊,防御,數量,銀兩,堅固度,等級,sm,會員) values ('"&b&"',"&idd&",'"&d&"','"&e&"',"&f&","&g&","&h&","&i&","&c&","&j&","&k&","&m&","&l&","&n&")"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示：操作成功！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
id=rs("id")
sql="update 物品 set 數量=數量+"&c&" where id="& id
conn.execute(sql)
rs.close
set rs=nothing
conn.close
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=#ffff00>【購買消息】</font>"& a&"向江湖總站購買了<font color=#ffffcc><b>"& b &"共"& c&"個</b></font>...大家恭喜他"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>
