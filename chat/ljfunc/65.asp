<%'����
function laobao(to1,toname)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ���� FROM �Τ� WHERE id="&info(9),conn
grade1=rs("����")
rs.close
rs.open "select �ʧO,�y�O,����,����  FROM �Τ� WHERE �m�W='"&toname&"'",conn
meimao=rs("�y�O") 
xingbie=rs("�ʧO") 
grade2=rs("����")
meipai=rs("����")
if grade1<garde2 then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('�H�a����g���A�״I�h�F�A�S��F�A�w�g�����F�I');}</script>" 
Response.End 
end if 
if meipai="�x��" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('�A�����x���H���i�榹�ާ@�I');}</script>" 
Response.End 
end if 
if xingbie="�k" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('�A�����k�H�ާ@�I');}</script>" 
Response.End 
end if 

'rs.close 
'%><!--#include file="jiu.asp"--><%
'sql="select * from �W�� where �m�W='"&toname&"'"
'set rs=connt.execute(sql)
'if not (rs.eof and rs.bof) then 
'Response.Write "<script language=javascript>{alert('�H�a�w�g�O�p�j�F�I');}</script>" 
'rs.close 
'set rs=nothing 
'conn.close 
'set conn=nothing 
'Response.End 
'end if 
'sql="insert into �W��(�m�W,������) values ('" & toname & "'," & meimao & ")"
'			connt.execute sql
'rs.close
'rs.open "select id FROM �Τ� WHERE id="&info(9),conn
'conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000,�D�w=�D�w-1000 where id="&info(9)
laobao=info(0)& "���L�֤l�u�O�n�ΡA" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>