<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
If Session("usepro") = true Then
'id=request("id")
win=request("win")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if win<>0 then
rs.open "select id from �Τ� where id="&info(9),conn
conn.execute "update �Τ� set ���O=���O+1000,��O=��O+1000,����=����+1000,���s=���s+1000,�Ȩ�=�Ȩ�+2000  where id="&info(9)
rs.close
rs.open "select ����,�����[,���s�[,���O�[,��O�[,���s,����,��O,���O from �Τ� where id="&info(9),conn
gjj=rs("�����[")
fyj=rs("���s�[")
nlj=(rs("����")*640+2000)+rs("���O�[")
tlj=(rs("����")*1500+5260)+rs("��O�[")
if rs("���s")>fyj then
conn.execute "update �Τ� set ���s=" & fyj & "  where id="&info(9)
end if
if rs("����")>gjj then
conn.execute "update �Τ� set ����=" & gjj & "  where id="&info(9)
end if
if rs("��O")>tlj then
conn.execute "update �Τ� set ��O=" & tlj & "  where id="&info(9)
end if
if rs("���O")>nlj then
conn.execute "update �Τ� set ���O=" & nlj & " where id="&info(9)
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Session("usepro")= false 
Response.Write "<script Language=Javascript>alert('�j�L�A�A�����رF���\�A���D�����O�B��O�B�����B���s�U1000�I�A�Ȩ�2000��I');</script>"
response.end

%>
<script language="vbscript">
window.close()
</script>	  
 <%
 else 

conn.execute "update �Τ� set ���O=���O-300,��O=��O-300,�Ȩ�=�Ȩ�-20 where id="&info(9)
Session("usepro")= false 
rs.open "select ���O,��O,���A from �Τ� where id="&info(9),conn
if rs("���O")<1 then conn.execute "update �Τ� set ���O=100 where id="&info(9)
if rs("��O")<0 or rs("���A")="��" then Response.Redirect "../exit.asp"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�泾�A���ɤ��m�\�A�{�b���ѤF�A�ٴ��Ө��ڡA�@�A��O�M���O300�A�Ȩ�20�I');</script>"
response.end
%>
<script language="vbscript">
window.close()
</script>	
<%
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Session("usepro")= false 
else 
response.write "�гq�L���`�~�|�ӥ��رF." 
response.end 
end if
%>
