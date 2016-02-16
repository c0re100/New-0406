<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if Session("minets")>0 then
	Session("minets")=Session("minets")-1
end if
if Session("minets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的礦石太少了能賣幾個錢啊？？！');</script>"
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩+"& Session("minets")*800 &" where id="&info(9)
Response.Write "<script Language=Javascript>alert('您采礦石："& session("minets") &"塊,賣了"& session("minets")*800 &"兩白銀！');</script>"
Session("minets")=0
Session("minesj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
