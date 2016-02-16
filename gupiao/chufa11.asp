<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "叫秈册ぱ秈布"
window.close()
</script>
<%Response.End
end if
if info(2)<10 then	Response.Redirect "../error.asp?id=439"
%>
<!--#include file="jhb.asp"-->
<%
Randomize
sj=int(rnd*26)+28
sql= "select * from 布 where sid="&sj        
set rs=conn.execute(sql)  
if (rs("讽玡基")-rs("秨絃基"))/rs("秨絃基")>=0.3 then 
ming=rs("穨")        
sql="update 布 set 讽玡基=讽玡基*0.8 where sid="&sj
conn.execute sql           
mess=ming&"布綝め┻扳布基20%"
elseif (rs("讽玡基")-rs("秨絃基"))/rs("秨絃基")<=-0.3 then
ming=rs("穨")
sql="update 布 set 讽玡基=讽玡基*1.2 where sid="&sj
conn.execute sql              
mess=ming&"布牟┏は紆基害20%"
else     
ming=rs("穨")     
Randomize
sj1=int(10*rnd)+1
if sj1>4 then
sql="update 布 set 讽玡基=讽玡基*"&(1-(sj1-1)/100)&" where sid="&sj
conn.execute sql 
mess=ming&"赋ㄆ蹿発禲布基禴"&(sj1-1)&"%"
else
sql="update 布 set 讽玡基=讽玡基*"&(1+sj1/100)&" where sid="&sj
conn.execute sql 
mess=ming&"ネ種砍订布基害"&sj1&"%"
end if
end if     
end if

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
	sd(194)=""
	sd(195)="產"
	sd(196)="FFFDAF"
	sd(197)="FFFDAF"
	sd(198)="癸"
	sd(199)="<font color=red>カ"& mess &"</font>"
	sd(200)=session("nowinroom")
	Application("ljjh_sd")=sd
	Application.UnLock
Response.Redirect "index.asp"

%>

