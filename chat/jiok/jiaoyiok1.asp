<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
grade=info(2)
myid=info(9)
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if application("jiutian_jiaoyi")="" then
Response.Write "<script Language=Javascript>alert('對不起，現在沒有拍賣，或者拍賣已經取消');window.close();</script>"
response.end
end if

sj=DateDiff("s",Application("jiutian_shijian"),now())
if sj<3 then
s=3-sj
Response.Write "<script language=JavaScript>{alert('提示：網絡在正處理數據，請您等上["&s&"秒鐘]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
Application.Lock
Application("jiutian_shijian")=now()
Application.UnLock


jian=split(application("jiutian_jiaoyi"),"|")
from1=jian(0)
wupin=jian(3)
jiage=int(jian(2))
jia=int(jiage*0.9)
sl=int(jian(4))
danwei=jian(5)
froo=jian(6)
if info(0)=froo then
Response.Write "<script Language=Javascript>alert('對不起，自己不能買自己的東西');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select 姓名 from 用戶 where id="&from1
rs.open sql,conn,1,1
nickk=rs("姓名")
rs.close
sql="SELECT 物品名,裝備 FROM 物品 WHERE (類型='暗器' or 類型='藥品' or 類型='法器' or 類型='法寶')  and 擁有者="&from1
rs.open sql,conn,1,1
wqm=rs("物品名")
my_wu=wupin
if my_wu=wqm and rs("裝備")=true then
Response.Write "<script Language=Javascript>alert('購買不成功，對方物品裝備著呢');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if

rs.close


sql="select 銀兩 from 用戶 where id="&info(9)
rs.open sql,conn,1,1

jinyan=int(rs("銀兩"))

if jiage>jinyan then
Response.Write "<script Language=Javascript>alert('購買不成功，你身上沒有這麼多的銀兩');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if

rs.close
sql="select * from 物品 where 物品名='" & wupin & "' and 擁有者="&from1
rs.open sql,conn,1,1
if rs.eof or rs.bof then
Response.Write "<script Language=Javascript>alert('購買不成功，對方沒有這樣的物品');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
wu=rs("物品名")
yin=rs("銀兩")
dj=rs("等級")
lx=rs("類型")
gj=rs("攻擊")
fy=rs("防御")
nl=rs("內力")
tl=rs("體力")
say=rs("說明")
sm=rs("sm")
rs.close
sql="select 數量 from 物品 where 物品名='" & wupin & "' and 擁有者="&info(9)
rs.open sql,conn,1,1
if rs.eof or rs.bof then
conn.execute"insert into 物品(物品名,擁有者,類型,說明,內力,體力,攻擊,防御,數量,銀兩,等級,sm,會員) values ('"&wu&"',"&info(9)&",'"&lx&"','"&say&"','"&nl&"','"&tl&"','"&gj&"','"&fy&"',"&sl&","&yin&","&dj&","&sm&","&aaao&")"
conn.execute"update 物品 set 數量=數量-'"&sl&"'  where 物品名='" & wupin & "' and 擁有者="&from1
if danwei="銀兩" then
conn.execute"update 用戶 set 銀兩=銀兩+'" & jia & "' where id="&from1
conn.execute"update 用戶 set 銀兩=銀兩-'" & jiage & "' where id="&info(9)
else
conn.execute"update 用戶 set 銀兩=銀兩+'" & jia & "' where id="&from1
conn.execute"update 用戶 set 銀兩=銀兩-'" & jiage & "' where id="&info(9)

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
msg="["&nickname&"]花了<font color=red>"&jiage&"</font>銀兩，從["&nickk&"]那買了<font color=FFFDAF>"&wupin&"</font>,交易成功，官府從中收取10%的稅。"
call jiaoyi()
Response.Write "<script Language=Javascript>alert('交易成功');window.close();</script>"
response.end
end if

shu=rs("數量")
if shu>20 then

Response.Write "<script Language=Javascript>alert('你身上的物品太多，不能再買!');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
else
conn.execute"update 物品 set 數量=數量+'"&sl&"'  where 物品名='" & wupin & "' and 擁有者="&info(9)
conn.execute"update 物品 set 數量=數量-'"&sl&"'  where 物品名='" & wupin & "' and 擁有者="&from1
if danwei="銀兩" then
conn.execute"update 用戶 set 銀兩=銀兩+'" & jia & "' where id="&from1
conn.execute"update 用戶 set 銀兩=銀兩-'" & jiage & "' where id="&info(9)
else
conn.execute"update 用戶 set 銀兩=銀兩+'" & jia & "' where id="&from1
conn.execute"update 用戶 set 銀兩=銀兩-'" & jiage & "' where id="&info(9)

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
msg="["&nickname&"]花了<font color=red>"&jiage&"</font>銀兩，從["&nickk&"]那買了<font color=FFFDAF>"&wupin&"</font>,交易成功，官府從中收取10%的稅。"
call jiaoyi()
Response.Write "<script Language=Javascript>alert('交易成功');window.close();</script>"
response.end
end if

sub jiaoyi()
application("jiutian_jiaoyi")=""
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
sd(199)="<font size=2><font color=FFFDAF>【拍賣成交】" &msg &"</font></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
%>