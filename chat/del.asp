<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');window.close();}</script>"
Response.End 
end if
if info(0)<>"�����`��" then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ӳo�̧@����A�Q���r�I');window.close();</script>"
	response.end
end if
if info(2)<7 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ӳo�̧@����A�Q���r�I');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���~�W from ���~ where id=" & id & " and �ƶq>0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�S���A�n�R�������~�I');history.go(-1);</script>"
	response.end
end if
conn.execute "delete * from ���~ where  id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ܡG�R�����~���\�I');history.go(-1);</script>"
	response.end
%>
 