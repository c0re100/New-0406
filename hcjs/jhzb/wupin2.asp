<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhzb")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if

id=lcase(trim(request("ID")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
active=lcase(trim(request("active")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,����,����,���s,�˳� FROM ���~ WHERE �֦���=" & info(9) & " and id=" & id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S���Ӫ��~�I�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
leixing=rs("����")
dj=rs("����")

zbgj=rs("����")
zbfy=rs("���s")
if  rs("�˳�")=true then

rs.close
	rs.open "SELECT ���� FROM �Τ� WHERE id="&info(9),conn
if dj>rs("����") then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�Ӹ˳ƻݭn"&dj&"���ׯ�˳ơI');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
select case leixing
case "�Y��"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='�Y��' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a1=0,b1=0 where id=" & info(9)
conn.execute "update ���~ set �˳�=false where id=" & id

case "����"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='����' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a2=0,b2=0 where id=" & info(9)
conn.execute "update ���~ set �˳�=false where id=" & id

case "�k��"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='�k��' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a3=0,b3=0 where id=" & info(9)
conn.execute "update ���~ set �˳�=false where id=" & id

case "���}"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='���}' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a4=0,b4=0 where id=" & info(9)
conn.execute "update ���~ set �˳�=false where id=" & id

case "�˹�"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='�˹�' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a5=0,b5=0 where id=" & info(9)
conn.execute "update ���~ set �˳�=false where id=" & id

case "����"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='����' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a6=0,b6=0 where id=" & info(9)
conn.execute "update ���~ set �˳�=false where id=" & id

case else
    rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ާ@���~�I�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end select
else
select case leixing
case "�Y��"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='�Y��' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a1="&zbgj&",b1="&zbfy&" where id=" & info(9)
conn.execute "update ���~ set �˳�=true where id=" & id

case "����"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='����' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a2="&zbgj&",b2="&zbfy&" where id=" & info(9)
conn.execute "update ���~ set �˳�=true where id=" & id

case "�k��"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='�k��' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a3="&zbgj&",b3="&zbfy&" where id=" & info(9)
conn.execute "update ���~ set �˳�=true where id=" & id

case "���}"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='���}' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a4="&zbgj&",b4="&zbfy&" where id=" & info(9)
conn.execute "update ���~ set �˳�=true where id=" & id

case "�˹�"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='�˹�' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a5="&zbgj&",b5="&zbfy&" where id=" & info(9)
conn.execute "update ���~ set �˳�=true where id=" & id

case "����"

rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and ����='����' and �˳�=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update ���~ set �˳�=false where id=" & idd
end if
conn.execute "update �Τ� set a6="&zbgj&",b6="&zbfy&" where id=" & info(9)
conn.execute "update ���~ set �˳�=true where id=" & id

case else
    rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ާ@���~�I�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end select
end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>location.href = 'head.asp';</script>"
	response.end
%>