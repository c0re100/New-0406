<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
pai=lcase(request("pai"))
a=request("a")
wg1=trim(request.form("wg1"))
id=request.form("id")
lx=trim(request.form("lx"))
nl=trim(request.form("nl"))
if lx<>"����" and  lx<>"���s" and lx<>"��_" then
Response.Write "<script language=JavaScript>{alert('�Z�\�������~�A�п�J�����B���s�B��_�I');location.href = 'setwg.asp';}</script>"
	response.end
end if

if instr(repass,"'")>0 or instr(pass,",")>0 or instr(nl,",")>0 or instr(wg,",")>0 or instr(wg1,",")>0 or instr(name,",")>0 then
	response.write "�A�n�r�I�o�O�T�a�A�ФŶ���!!!!"
	response.end
end if
if instr(wg,"or")<>0 or instr(pai,"or")<>0 or instr(wg1,"or")<>0 or instr(nl,"or")<>0 or instr(name,"or")<>0 or instr(pass,"or")<>0 then
	Response.Write "<script language=JavaScript>{alert('�A�n�r�I�o�O�T�a�A�ФŶ���!!!!');location.href = 'setwg.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �x�� FROM ���� where trim(�x��)='" & info(0) & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�n�r�I�o�O�T�a�A�ФŶ���!!!!');location.href = 'setwg.asp';}</script>"
	response.end
end if

if a="m" then
if Request.Form("submit")="�R��" then
	wgid=int(abs(request("id")))
	conn.Execute "delete * from �Z�\  where id="&wgid
end if

	wgid=int(abs(request("id")))
	conn.Execute "update �Z�\ set �Z�\='" & wg1 & "',����='" & lx & "', ���O=" & nl & " where id="&wgid
end if
if a="n" then
rs.close
rs.open "SELECT �Z�\ FROM �Z�\ where trim(�Z�\)='" & wg1 & "'",conn
if not rs.eof or not rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ӪZ�\�ۦ�����̤w�g���F�A�Э��ب�L�ۦ�!!!!');location.href = 'setwg.asp';}</script>"
	response.end
end if
	conn.Execute "insert into �Z�\(�Z�\,����,�ŧO,����,���O) values ('" & wg1 & "','" & lx & "',1,'" & pai & "'," & nl & ")"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�ާ@���\�I');location.href = 'setwg.asp';}</script>"
	Response.End
%>
<body background="../linjianww/0064.jpg">
