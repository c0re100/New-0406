<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)<>"�����`��"  then  response.redirect "../error.asp?id=425"

shopname=Request.Form("shopname")
name=Request.Form("name")
memo=Request.Form("memo")
shoptype=Request.Form("shoptype")
shopmoney=Request.Form("shopmoney")

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
if (not rs1.EOF) or (not rs1.BOF) then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("���ө��W�w�Q�`�U�Э��s���!")
history.back(1)
</script>
<%Response.End 
end if

sql="insert into �ө� (�ө��W,���D,id,����,�ө�����,�`�U���) values ('"&shopname&"','"&name&"',"&id&",'"&memo&"','"&shoptype&"',"&shopmoney&")"
conn1.Execute(sql)
set rs1=nothing
conn1.close
set conn1=nothing
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�����j����"&name&"��["&shopname&"]�ө��}�i</font>�`�U���"&shopmoney&"....<font color=FFD7F4>�w��j�a���{�I</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
</center>
<script language="vbscript">
alert("�K�[���\!")
window.close()
</script>