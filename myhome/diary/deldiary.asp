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

 
<!--#include file="../data2.asp"-->

<%
rs.open "select * from diarymessage where id="&request.form("id"),conn,1,3
if rs.bof then

%>

<script language="Vbscript">
msgbox "記錄不存在！",0,"提示"
history.back
</script>

<%
rs.close
conn.close
Response.End
end if
rs.delete
rs.close
conn.close

%>


<script language="Vbscript">
msgbox "成功刪除！",0,"提示"
history.back
</script>