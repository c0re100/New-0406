<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=440"
id=request("id")
my=info(0)
%><!--#include file="dadata.asp"-->
<%
sql="select �t�� from �Τ� where �m�W='" & my & "'  and �t��<>'�L' and �t��<>'���w'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('�A�S���t���A����q�ʦ��s�b');history.back();</script>"
response.end
else
peiou=trim(rs("�t��"))
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
sql="update �Τ� set �Ȩ�=�Ȩ�-" & yin & " where id="&info(9)
rs=conn.execute(sql)%>
			
<%
sql="select �b�|�W from �b�|�C�� where �b�|�W='" & wu & "' and �֦���='" & my & "'"
set rs=connt.execute(sql)
if rs.eof or rs.bof then
sql="insert into �b�|�C��(�b�|�W,�֦���,����,���O,��O,����,�ƶq,���,�ѥ[��,�ɶ�) values ('"&wu&"','"&my&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",'"&peiou&"',now())"
rs=connt.execute(sql)
connt.close
			
set connt=nothing
if Instr(Application("ljjh_useronlinename"&session("nowinroom"))," "&peiou&" ")<>0 then

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
sd(195)=peiou
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i�n�����j"&my&"�b����j�s����"&wu&"�U��"&peiou&"�|��<font color=red>��"&lx&"��</font>�b�|�A"&peiou&"�֥h�r�A����..</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end if			
Response.Write "<script language=javascript>alert('���ߧA�A�A���s�b�w�ন�\�A�h�i�D�A���L(�o)�a�I');history.back();</script>"			
Response.Redirect "jd.asp"
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
%>