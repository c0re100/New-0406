<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �Τ� WHERE ����<>'�x��' order by mvalue desc"
Set Rs=conn.Execute(sql)
if rs.eof and rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ثe�S����n���Ʀ�I�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
else
do while not rs.eof
conn.execute "update �Τ� set �k�O=�k�O+50 where �m�W='"&rs("�m�W")&"'"
rs.movenext
file=file+1
if file>20 then Exit Do
loop
end if
set rs=nothing
conn.close
set conn=nothing
mess="��n���Ʀ�e20�W���y�k�O50���A�j�a���P�A�~��V�O�r�I"
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
sd(199)="<font size=3 color=red>�i���y�����j</font><font color=FFD7F4 size=3>"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('�{�b�w�N�Ҧ���n���e20�W�����a���y�F50�I�k�O�I�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>