<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

name=Request.Form ("name")
job=Request.Form ("job")
shopname=Request.Form ("shopname")

sql="select ¾�~ from �Τ� where �m�W='" & name & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%>
<script language="vbscript">
alert("�S�����H�I")
history.back(1)
</script>
<%
end if
if rs("¾�~")="�t�Įv" then 
%>
<script language="vbscript">
alert("���H�w�O�t�Įv�I")
history.back(1)
</script>
<%
end if
sql="update �Τ� set ¾�~='"& job &"',����='"&shopname&"' where �m�W='" & name & "'"
conn.execute sql
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�z��"&name&"�ۦ����A���G"& job &"�I���T�w��^�I');</script>"
%>
<script language="vbscript">
history.back(1)
</script>