<%
function check(key,n)
if trim(key)="" then
%>
<script language=javascript>
window.alert("<%="�п�J"&n&"����!"%>");
history.back();
</script>
<%
end if
if instr(key,"<")<>0 or instr(key,">")<>0 or instr(key,",")<>0 or instr(key,"'")<>0 or  instr(key,chr(34))<>0 or instr(key," ")<>0 or instr(key,"%")<>0 or instr(key,"&")<>0 or instr(key,"\")<>0 or instr(key,"/")<>0 then
%>
<script language=javascript>
window.alert("<%=n&"����J�����D�k�r��!"%>");
history.back();
</script>
<%
else
check=key
end if
end function
%>