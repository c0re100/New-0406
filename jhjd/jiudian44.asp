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
if info(0)="" then Response.Redirect "../error.asp?id=440"
id=trim(request.form("id"))
qingren=trim(request.form("name"))
my=info(0)
%><!--#include file="dadata.asp"-->
<%
sql="select * from �Τ� where id="&info(9)
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "�A���O���򤤤H�A����q�ʰs�b"
response.end
else
sql="select * from �Τ� where �m�W='" & qingren & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('���򤤨S���A��J���W�r!');history.back();</script>"
response.end
else
qingren=rs("�m�W")
sql="SELECT * FROM �b�| where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("�b�|�W")
yin=rs("���")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
jb=rs("����")
sl=rs("�ƶq")%>

<%
sql="select * from �Τ� where id="&info(9)
rs=conn.execute(sql)
if yin<=rs("�Ȩ�") then
sql="update �Τ� set �Ȩ�=�Ȩ�-" & yin & " where id="&info(9)
rs=conn.execute(sql)%>
			
<%
sql="select * from �b�|�C�� where �b�|�W='" & wu & "' and �֦���='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into �b�|�C��(�b�|�W,�֦���,����,���O,��O,����,�ƶq,���,�ѥ[��,�ɶ�) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&qingren&"',now())"
rs=connt.execute(sql)
connt.close
set connt=nothing
if Instr(Application("ljjh_useronlinename")," "&qingren&" ")<>0 then

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
sd(195)=qingren
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i�n�����j"&my&"�b����j�s����"&wu&"�U��"&qingren&"�|��<font color=red>��"&lx&"��</font>�b�|�A"&qingren&"�֥h�r�A����..</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock


end if
			
Response.Write "<script language=javascript>alert('���ߡA�A��"&qingren&"�w���s�b�w�g�ǳƦn�F�I');window.close();</script>"
response.end
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('����w�s�b�A��]�G�A�w�w�F�P�˳W�檺�s�u�I');history.back();</script>"
response.end
				
end if
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('����w�s�b�A��]�G�A���Ȩ⤣���F');history.back();</script>"
response.end
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
%>