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

 
<!--#include file="../data1.asp"-->
<%rs.open "select id from bookmark where id="&request.querystring("id")&"",conn,1,1
if rs.bof then%>
<script language="Vbscript">
msgbox"沒有此紀錄！",0,"提示"
history.back
</script>
<%
rs.close
Response.End
end if%>

<%number=rs("number")
conn.execute "delete * from bookmark where id="&request.querystring("id")&""
rs.close
rs.open "select count,allcount from bookmark_count where user='"&request.querystring("user")&"'",conn,1,3
if number>rs("count") then
numbernow=0
else
numbernow=rs("count")-number
rs.update "count",numbernow
end if
if rs("allcount")>0 then
rs.update "allcount",rs("allcount")-1
else
rs.update "allcount",0
end if
rs.close
conn.close
%>
<script language="Vbscript">
msgbox"您已經成功地刪除了書簽！",0,"提示"
history.back
</script>