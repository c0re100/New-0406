<%@ LANGUAGE=VBScript%>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from ���~ where �ƶq>0 and ���~�W='�K�o�d' and �֦���="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing	
	Response.Write "<script language=JavaScript>{alert('�A�K�o�d�ܡH�I');location.href = 'listlao.asp';}</script>"
	Response.End
end if
conn.Execute "update ���~ set �ƶq=�ƶq-1 where �֦���="&info(9)&" and ���~�W='�K�o�d'"
conn.execute "update �Τ� set ���A='���`',�n��=now() where �m�W='" & name & "'"
msg="<font color=green>�i�K�o�d�j</font><font color=ffffff>"&info(0)&"</font>�����ϥΤF�K�o�d��b�c�̪�<font color=blue>" & name & "</font>���O���F�X�ӡI"
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
sd(196)="660099"
sd(197)="660099"
sd(198)="��"
sd(199)="<font color=red>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "listlao.asp"
%>