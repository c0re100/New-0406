<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from �Τ� where �m�W='" & name & "' and lastkick='" & my & "'",conn 
if rs.eof or rs.bof then
rs.close
conn.close
set rs=nothing	
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�٨S���F�o�ӤH�O�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select ���H�Ī� from ���� where �m�W2='" & my & "' and �Q����='" & name & "'" ,conn
shangjin=rs("���H�Ī�")
conn.execute "delete * from ���� where �m�W2='" & my & "' and �Q����='" & name & "'"
conn.execute "update �Τ� set lastkick='�L' where �m�W='"&name&"'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&shangjin&" where id="&info(9)
rs.close
conn.close
set rs=nothing	
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ߧA�o��A���o���@���Ȩ�["&shangjin&"]');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>