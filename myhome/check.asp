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