<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then  response.redirect "../error.asp?id=425"

shopname=Request.Form("shopname")
name=Request.Form("name")
memo=Request.Form("memo")
shoptype=Request.Form("shoptype")
shopmoney=Request.Form("shopmoney")
id=Request.Form("id")
if shopname="" or name="" or shopmoney="" or memo="" or shoptype="�п��" then 
%><script language="vbscript">
alert("�C������Ъ�^����!")
history.back(1)
</script>
<%Response.End 
end if
%><!--#include file="data.asp"--><%
sql="select �s��,id from �Τ� where �m�W='"&name&"'"
set rs=conn.Execute (sql)
if  rs.EOF or  rs.BOF then
set rs=nothing
conn.close
set conn=nothing%>
<script language="vbscript">
alert("����̨S���o�ӤH�r!")
history.back(1)
</script>
<%Response.End 
end if
id=rs("id")
if rs("�s��")<int(shopmoney) then
%><script language="vbscript">
alert("�`�U����Ӧh�A�s�ڤӤ�!")
history.back(1)
</script>
<%Response.End 
end if
sql="SELECT * FROM �ө� where �ө��W='"&shopname&"'"
set rs1=conn1.Execute (sql)
sql="update �ө� set ���D='"&name&"',id="&id&",����='"&memo&"',�ө�����='"&shoptype&"',�`�U���="&shopmoney&" where  �ө��W='"&shopname&"'"
conn1.Execute (sql)
set rs1=nothing
conn1.close
set conn1=nothing
%>
</center>
<script language="vbscript">
alert("��令�\�I")
window.close()
</script>