<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" or info(5)="" then Response.Redirect "../error.asp?id=440"
id=request("id")
my=info(0)
mypai=info(5)
%><!--#include file="dadata.asp"-->
<%
sql="select ID from �Τ� where id="&info(9)&" and ����<>'�L' and ����<>'�C�L'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('�A�S�������A����q�ʦ��s�b');history.back();</script>"
response.end
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
if info(5)="����" or info(5)="�Ư���" then
mypai="�x��"
end if
sql="select �Ȩ� from �Τ� where id="&info(9)&" and ����='"&mypai&"'"
rs=conn.execute(sql)
if yin<=rs("�Ȩ�") then
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & " where id="&info(9)
%>
			
<%
sql="select �b�|�W from �b�|�C�� where �b�|�W='" & wu & "' and �֦���='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into �b�|�C��(�b�|�W,�֦���,����,���O,��O,����,�ƶq,���,����,�ɶ�) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&mypai&"',now())"
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
sd(199)="<font color=FFFDAF>�i�n�����j"&my&"�b����j�s����"&wu&"�U��"&mypai&"�|��<font color=red>��"&lx&"��</font>�b�|�A"&mypai&"�����֥h�r�A�ߤF�N�S����F�C�C�C</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

Response.Write "<script language=javascript>alert('���ߧA�A�A���s�b�w�ন�\�A�h�i�D�A���B�ͧa�I');history.back();</script>"			
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
	set rs=nothing
	conn.close
	set conn=nothing
end if
%>