<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from �Τ� where id="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���O���򤤤H�A�u�I');location.href = 'card.asp';}</script>"
	Response.End
end if
rs.close
rs.open "SELECT ���~�W,�Ȩ�,����,����,���s,��O,���O,��T��,����,sm FROM ���~�R where ID=" & id & " and ����='�d��'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�S���A�n�ʶR���|���d���H!');location.href = 'card.asp';}</script>"
	response.end
end if
card=rs("���~�W")
hyyin=rs("�Ȩ�")
hyyin1=rs("����")
say=rs("����")
sm=rs("sm")
fy=rs("���s")
tl=rs("��O")
nl=rs("���O")
jgd=rs("��T��")
dj=rs("����")
if hyyin>=1 then
rs.close
rs.open "select �|���O from �Τ� where id="&info(9)
if hyyin>rs("�|���O") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���|���O�����I!');location.href = 'card.asp';}</script>"
	response.end
end if
conn.execute "update �Τ� set �|���O=�|���O-" &hyyin & " where id="&info(9)
rs.close
rs.open "select �|���O from �Τ� where id="&info(9)
if rs("�|���O")<0 then
conn.execute "update �Τ� set �|���O=0 where id="&info(9)
end if
else
rs.close
rs.open "select ���� from �Τ� where id="&info(9)
if hyyin1>rs("����") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�����������I!');location.href = 'card.asp';}</script>"
	response.end
end if
conn.execute "update �Τ� set ����=����-" &hyyin1 & " where id="&info(9)
rs.close
rs.open "select ���� from �Τ� where id="&info(9)
if rs("����")<0 then
conn.execute "update �Τ� set ����=0 where id="&info(9)
end if
end if
rs.close
rs.open "select id,�ƶq from ���~ where ����='�d��' and ���~�W='" & card & "' and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~ (���~�W,�֦���,����,����,���s,���O,��O,��T��,�Ȩ�,����,�ƶq,sm,����,�|��) values ('"&card&"',"&info(9)&",'�d��',"&hyyin1&","&fy&","&nl&","&tl&","&jgd&","&hyyin&",'"&say&"',1,"&sm&","&dj&","&aaao&")"
else
	id=rs("id")
	sl=rs("�ƶq")
	conn.execute "update ���~ set �Ȩ�=" & hyyin & ",�ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���d��["&card&"]�ʶR���\,�{��"&sl+1&"�ӡI');location.href = 'card.asp';}</script>"
response.end
%>
