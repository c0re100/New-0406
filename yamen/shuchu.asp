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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
my=info(0)
id=request("id")
'����Τ�
rs.open "select grade,ID from �Τ� where id="&info(9),conn
If Rs.Bof OR Rs.Eof Then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�A���O���򤤤H�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if  rs("grade")<6 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�L���v�O�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
UID=rs("ID")
rs.close
rs.open "select id from �Τ� where ���A='�c' or ���A='��'  and id=" & id
if  rs.eof and  rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�S�o�ӤH�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
conn.execute "update �Τ� set ���A='���`',�n��=now() where id=" & id
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
sd(198)="�i����j"
'sd(199)="<font color=red>�i����j" & rs("�m�W") & "�Q�L�o����(����x["&info(0)&"])</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���񦨥\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
