<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=440"
id=trim(request.form("id"))
qingren=trim(request.form("name"))
my=info(0)
%><!--#include file="dadata.asp"-->
<%
sql="select * from 用戶 where id="&info(9)
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "你不是江湖中人，不能訂購酒宴"
response.end
else
sql="select * from 用戶 where 姓名='" & qingren & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('江湖中沒有你輸入的名字!');history.back();</script>"
response.end
else
qingren=rs("姓名")
sql="SELECT * FROM 宴會 where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("宴會名")
yin=rs("售價")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
jb=rs("金幣")
sl=rs("數量")%>

<%
sql="select * from 用戶 where id="&info(9)
rs=conn.execute(sql)
if yin<=rs("銀兩") then
sql="update 用戶 set 銀兩=銀兩-" & yin & " where id="&info(9)
rs=conn.execute(sql)%>
			
<%
sql="select * from 宴會列表 where 宴會名='" & wu & "' and 擁有者='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into 宴會列表(宴會名,擁有者,類型,內力,體力,金幣,數量,售價,參加者,時間) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&qingren&"',now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
if Instr(Application("ljjh_useronlinename")," "&qingren&" ")<>0 then

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
sd(195)=qingren
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=FFFDAF>【好消息】"&my&"在江湖大酒店的"&wu&"廳為"&qingren&"舉行<font color=red>※"&lx&"※</font>宴會，"&qingren&"快去呀，呵呵..</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock


end if
			
Response.Write "<script language=javascript>alert('恭喜，你為"&qingren&"定的酒宴已經準備好了！');window.close();</script>"
response.end
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('不能定酒宴，原因：你已定了同樣規格的酒席！');history.back();</script>"
response.end
				
end if
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('不能定酒宴，原因：你的銀兩不夠了');history.back();</script>"
response.end
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
%>