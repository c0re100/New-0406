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
%><!--#include file="dadata.asp"-->
<%
rs.open "select * from �Τ� where id="&info(9),conn
if rs.eof or rs.bof then
	response.write "�A���O���򤤤H�A����q�ʰs�b"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
rs.open "SELECT * FROM �b�| where ID=" & id,conn
wu=rs("�b�|�W")
yin=rs("���")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
jb=rs("����")
sl=rs("�ƶq")%>
<%
rs.close
rs.open "select * from �Τ� where id="&info(9),conn
if yin>=rs("�Ȩ�") then
	rs.close
	set rs=nothing
	connt.close
	set connt=nothing
	conn.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('����w�s�b�A��]�G�A���Ȩ⤣���F');history.back();</script>"
	response.end
end if
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & " where �m�W='" & info(0) & "'"
	rs.close
	rs.open "select * from �b�|�C�� where �b�|�W='" & wu & "' and �֦���='" & info(0) & "'",conn
if rs.eof or rs.bof then
	connt.execute "insert into �b�|�C��(�b�|�W,�֦���,����,���O,��O,����,�ƶq,���,�ɶ�) values ('"&wu&"','"&info(0)&"','"&lx&"',"&nl&","&tl&","&jb&","&sl&","&yin&",now())"
	rs.close
	set rs=nothing
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
	sd(199)="<font color=FFFDAF>�i�n�����j"&info(0)&"�b����j�s����"&wu&"�U�|��<font color=red>��"&lx&"��</font>�b�|�A�j�a���֥h�r�A�ߤF�N�S����F�C�C�C</font>"
	sd(200)=session("nowinroom")
	Application("ljjh_sd")=sd
	Application.UnLock
	Response.Redirect "jd.asp"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	connt.close
	set connt=nothing
	Response.Write "<script language=javascript>alert('����w�s�b�A��]�G�A�w�w�F�P�˳W�檺�s�u�I');history.back();</script>"
	response.end
end if
%>
