<%
if session("jiutian_name")="" or session("jiutian_id")="" then 
Response.Write "<script Language=Javascript>alert('�藍�_,�A�|���n���A�Τw�g�W���_�}�F�s���A�Э��s�n���I');window.close();</script>"
 response.end
end if
myid=session("jiutian_id")

%>