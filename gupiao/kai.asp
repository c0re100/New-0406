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
sql= "select * from 布 where sid=28"        
set rs=conn.execute(sql)
if rs("篈")="秨" then 
set rs=nothing
Response.Write "<script Language=Javascript>alert('ヘ玡カ竒琌秨絃篈');location.href = 'javascript:history.go(-1)';</script>"
else
for i=28 to 55  
sql= "select * from 布 where sid="&i        
set rs=conn.execute(sql)   
sql="update 布 set 秨絃基='"&rs("讽玡基")&"',篈='秨',ら戳='"&date()&"'where sid="&i
conn.execute sql
next

end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "index.asp"

%>

