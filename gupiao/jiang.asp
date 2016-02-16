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
if  info(2)<10 then
	Response.Redirect "../error.asp?id=439"
else%>
<!--#include file="jhb.asp"-->
<%
for sj=28 to 55
sql= "select * from 布 where sid="&sj        
set rs=conn.execute(sql) 
if (rs("讽玡基")-rs("秨絃基"))/rs("秨絃基")>-0.01 then 
sql="update 布 set 讽玡基=讽玡基*0.9 where sid="&sj
conn.execute sql 
end if
next 
mess="布カ初ǐ旱カ┮Τ布癵ぃ基10%"
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
       set rs=nothing
conn.close
set conn=nothing
Response.Redirect "index.asp"
end if
%>

