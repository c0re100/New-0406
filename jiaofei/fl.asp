<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	response.end
end if
if Application("ljjh_jiaofei")<>"�g��" then
	Response.Write "<script Language=Javascript>alert('���ܡG�{�b�٨S���g��H�I');window.close();</script>"
	response.end
end if
'if info(2)>=7  then
'	Response.Write "<script Language=Javascript>alert('���ܡG�x�����H���ζϭ�I�I');window.close();</script>"
'	response.end
'end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���O,��O,�Ȩ� from �Τ� where id="&info(9),conn
if rs("���O")<10000 or rs("��O")<10000 or rs("�Ȩ�")<100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ϭ�ݭn���B��O�U10000�A�ջ�100000��I�I');window.close();</script>"
	response.end
end if
session("dalie")=true
conn.execute "update �Τ� set ���O=���O-10000,��O=��O-10000,�Ȩ�=�Ȩ�-100000 where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script>
function dalie()
{
location.href='dalie.asp';
}
</script>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="444" height="278">
<param name=movie value="dalie.swf">
<param name=quality value=high>
<embed src="dalie.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
</embed>
</object>
