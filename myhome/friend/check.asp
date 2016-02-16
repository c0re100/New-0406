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

<%
function check(key,n)
if trim(key)="" then
%>
<script language=javascript>
window.alert("<%="請輸入"&n&"的值!"%>");
history.back();
</script>
<%
end if
if instr(key,"<")<>0 or instr(key,">")<>0 or instr(key,",")<>0 or instr(key,"'")<>0 or  instr(key,chr(34))<>0 or instr(key," ")<>0 or instr(key,"%")<>0 or instr(key,"&")<>0 or instr(key,"\")<>0 or instr(key,"/")<>0 then
%>
<script language=javascript>
window.alert("<%=n&"的輸入中有非法字符!"%>");
history.back();
</script>
<%
else
check=key
end if
end function
%>

<html><script language="JavaScript">                                                                  </script></html>