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
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if
if info(2)<10 then	Response.Redirect "../error.asp?id=439"
%>
<!--#include file="jhb.asp"-->
<%
sql= "select * from �Ѳ� where sid=28"        
set rs=conn.execute(sql)
if rs("���A")="�}" then 
set rs=nothing
Response.Write "<script Language=Javascript>alert('�ثe�ѥ��w�g�O�}�L���A�F�I');location.href = 'javascript:history.go(-1)';</script>"
else
for i=28 to 55  
sql= "select * from �Ѳ� where sid="&i        
set rs=conn.execute(sql)   
sql="update �Ѳ� set �}�L����='"&rs("��e����")&"',���A='�}',���='"&date()&"'where sid="&i
conn.execute sql
next

end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "index.asp"

%>

