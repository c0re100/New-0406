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
rs.open "select �Ȩ�,����,����,�ʧO,�D�w from �Τ� where id="&info(9),conn
yl=rs("�Ȩ�")
mp=rs("����")
sf=rs("����")
xb=rs("�ʧO")
if yl<200000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="�A���������A�S��k�F�A�x�����s�W�w���A�^�h�n�n�Q�Q��k�A�ӧa�I"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"

response.end
end if
if rs("�ʧO")="�k" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="�o�̬O�M�|�q�r�A�A�����աI"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"

response.end
end if
if mp="�X�a" then
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&yl&",����='�C�L',����='�̤l' where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="���ߡA�A�w�g�٫U�F�A�ۤv�����ۤv���a�аh�X�᭫�s�n��"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"
response.end
end if
if mp="�x��" or sf="�x��"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
mess="�A�O�x���άO�x������X�a�I"
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"
response.end
end if
yl1=yl-200000
dd=rs("�D�w")-300
mess="���ߡA�A�w�g���\�X�a�F�A�H�᪺���ۤv���n�a�б��X�᭫�s�n��"
conn.execute "update �Τ� set �Ȩ�='"&yl1&"',�D�w='"&dd&"',����='�X�a',����='�M�|' where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');window.close();}</script>"
response.end%>