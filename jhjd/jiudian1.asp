<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=request("id")
my=info(0)
%><!--#include file="dadata.asp"-->
<%
sql="select ID from �Τ� where id="&info(9)
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�A�A���O���򤤤H');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
else

sql="SELECT �b�|�W,���,����,���O,��O,����,�ƶq FROM �b�| where ID=" & id
Set Rs=connt.Execute(sql)
wu=rs("�b�|�W")
yin=rs("���")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
jb=rs("����")
sl=rs("�ƶq")%>

<%
sql="select �Ȩ� from �Τ� where id="&info(9)
rs=conn.execute(sql)
if yin<=rs("�Ȩ�") then
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & " where id="&info(9)
%>
			
<%
sql="select �b�|�W from �b�|�C�� where �b�|�W='" & wu & "' and �֦���='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into �b�|�C��(�b�|�W,�֦���,����,���O,��O,����,�ƶq,���,�ɶ�) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
rs=connt.execute(sql)
connt.close
set connt=nothing

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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i�n�����j"&my&"�b����j�s����"&wu&"�U�|��<font color=red>��"&lx&"��</font>�b�|�A�j�a���֥h�r�A�ߤF�N�S����F�C�C�C</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

			
Response.Redirect "jd.asp"
else
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('����w�s�b�A��]�G�A�w�w�F�P�˳W�檺�s�u�I');history.back();</script>"
response.end
				
end if
else
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('����w�s�b�A��]�G�A���Ȩ⤣���F');history.back();</script>"
response.end
end if
end if
	set rs=nothing
	conn.close
	set conn=nothing
%>
