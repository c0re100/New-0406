<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
	sex=request.form("sex")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �Ȩ�,���,��O FROM �Τ� where id="&info(9)&" and �ʧO= '" &sex&"'",conn
if rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�A�A���O���򤤤H');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if
if rs("�Ȩ�")<380 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�藍�_�A�A��������');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if
if day(rs("���"))<day(now()) or month(rs("���"))<month(now()) or year(rs("���"))<year(now()) or isnull(rs("���")) then
energy=rs("��O")
tempdate=date
energytemp=energy+1000
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-300,���='"&tempdate&"',��O='"&energytemp&"' where id="&info(9)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�~���F�Ŭu�D�Pı���Q�h�F�C');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
else
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�藍�_�A�A�w�g�~�L�F');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if
%>
