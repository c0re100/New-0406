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
if Session("icets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你一桶冰也沒有采，你拿什麼賣錢？？？');</script>"
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩+"& Session("icets")*2000 &" where id="&info(9)

Response.Write "<script Language=Javascript>alert('您采冰："& session("icets") &"桶,賣了"& session("icets")*2000 &"兩白銀！');</script>"
Session("icets")=0
Session("icejl1")=0
Session("cbsj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
