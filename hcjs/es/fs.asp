<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=210"
sl=abs(int(Request.form("wpsl")))
fs=int(Request.form("wpfs"))
id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ƶq,���~�W from ���~ where id=" & id & " and �ƶq>0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S���o�ت��~�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if int(rs("�ƶq"))<sl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ާ@���~�A["&rs("���~�W")&"]�u��"&rs("�ƶq")&"�Ӫ��~�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
select case fs
case 1
	Response.Redirect "zs.asp?id="& id &"&sl="& sl
case 2
	Response.Redirect "jy.asp?id="& id &"&sl="& sl
case 3
	Response.Redirect "es.asp?id="& id &"&sl="& sl
case 4
	Response.Redirect "bxg.asp?id="& id &"&sl="& sl
end select
%>