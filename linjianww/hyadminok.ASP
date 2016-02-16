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
if info(2)<10   then Response.Redirect "../error.asp?id=900"
username=Request.Form("search")
sjyear=Request.form("year")
sjmonth=Request.form("month")
sjday=Request.form("day")
hygrade=int(Request.form("hygrade"))
hylx=int(Request.form("hylx"))
hydate=sjyear&"-"&sjmonth&"-"&sjday
select case hygrade
case 1
jhgrade=1
hyvalue=1250
hyjia=0
case 2
jhgrade=1
hyvalue=5000
hyjia=100
case 3
jhgrade=1
hyvalue=20000
hyjia=200
case 4
jhgrade=1
hyvalue=35000
hyjia=300
case 5
jhgrade=1
hyvalue=45000
hyjia=400
case 6
jhgrade=1
hyvalue=55000
hyjia=500
case 7
jhgrade=1
hyvalue=65000
hyjia=600
case 8
jhgrade=1
hyvalue=75000
hyjia=700
case 9
jhgrade=1
hyvalue=85000
hyjia=800
case 10
jhgrade=1
hyvalue=105000
hyjia=1000
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
Set rs=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM 用戶 where 姓名='"&username&"'"
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
	Response.Write "<script language=javascript>alert('抱歉你所要查找的人我們找不到！請查看是否正確！');history.back();</script>"
Response.End
else
if  rs("grade")>5 then
Response.Write "<script language=javascript>alert('抱歉["&username&"]是官府人員，請手動修改！');history.back();</script>"
Response.End
end if
	if rs("會員等級")=0 and rs("grade")<6 then
		sql="update 用戶 set 會員等級="& hygrade &",會員日期='"&hydate&"',內力加=內力加+"& hyjia &",體力加=體力加+"& hyjia &",武功加=武功加+"& hyjia &",攻擊加=攻擊加+"& hyjia &",防御加=防御加+"& hyjia &",知質加=知質加+"& hyjia &",grade="& jhgrade &",allvalue=allvalue+"& hyvalue &" where 姓名='"&username&"'"
		Set rs=conn.Execute(sql)
	else
		Response.Write "<script language=javascript>alert('抱歉["&username&"]現在已經是["&rs("會員等級")&"級]江湖會員，請手動更改！');history.back();</script>"
Response.End
	end if
end if
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','會員添加')"
conn.Execute(sql)
set rs=nothing
conn.Close
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
if hylx=1 then
sd(199)="<font color=FFD7F4>【會員消息】歡迎新會員["&username&"]加入會員行列!</font>" 
else
sd(199)="<font color=FFD7F4>【會員消息】歡迎新會員(月排前2名)["&username&"]加入會員行列!</font>"
end if
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=javascript>alert('恭喜你的江湖["& hygrade &"]級會員設置完成！');history.back();</script>"
Response.End
%> 