<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ�,����,�ʧO,�D�w from �Τ� where id="&info(9),conn
yl=rs("�Ȩ�")
mp=rs("����")
xb=rs("�ʧO")
if yl<200000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="�A���������A�ѹꪺ�X�a�a�A�ۤv�����i�O�ۤv�諸�@�I"
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'huansu.asp';}</script>"
response.end
end if
if mp<>"�X�a" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="�A�����O���򪺩M�|�Υ��h�A���̨Ӫ��y�]�A�ڭ̤��w��A�@�A�U�s�a�I"
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'huansu.asp';}</script>"
response.end
end if
yl1=yl-200000
dd=rs("�D�w")-300
mess="���ߡA�A�w�g�٫U�F�A�ۤv�����ۤv���a�аh�X�᭫�s�n��"
if xb="�k" then
conn.execute "update �Τ� set �Ȩ�='"&yl1&"',�D�w='"&dd&"',����='�C�L',����='�L' where id="&info(9)
else
conn.execute "update �Τ� set �Ȩ�='"&yl1&"',�D�w='"&dd&"',����='�C�L',����='�L' where id="&info(9)
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = '../exit.asp';}</script>"
response.end%>
<html>