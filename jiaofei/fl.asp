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
	Response.Write "<script Language=Javascript>alert('提示：你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	response.end
end if
if Application("ljjh_jiaofei")<>"土匪" then
	Response.Write "<script Language=Javascript>alert('提示：現在還沒有土匪？！');window.close();</script>"
	response.end
end if
'if info(2)>=7  then
'	Response.Write "<script Language=Javascript>alert('提示：官府中人不用剿匪！！');window.close();</script>"
'	response.end
'end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 內力,體力,銀兩 from 用戶 where id="&info(9),conn
if rs("內力")<10000 or rs("體力")<10000 or rs("銀兩")<100000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：剿匪需要內、體力各10000，白銀100000兩！！');window.close();</script>"
	response.end
end if
session("dalie")=true
conn.execute "update 用戶 set 內力=內力-10000,體力=體力-10000,銀兩=銀兩-100000 where id="&info(9)
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
