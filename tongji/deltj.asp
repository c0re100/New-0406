<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10 then Response.Redirect "../error.asp?id=461" end if%>
<!--#include file="data1.asp"-->
<%
tjname=request.form("tjname")
sql="delete * from �q�r�O WHERE �q�r�H��='"& Request("tjname") &"'"
	conn.execute sql
conn.close
tj="<font color=red>�q�r�H��["& Request("tjname") &"]��������A����[" & info(0) & "]���\�Ѱ��q�r�I</font>"
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
sd(199)="<font color=FFD7F4>�i�Ѱ��q�r�j</font>"&tj
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language=vbscript> 
MsgBox "�����G�w�g�R���ӳq�r�O�I" 
location.href = "javascript:history.back()" 
</script> 