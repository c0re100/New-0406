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
	Response.Write "<script Language=Javascript>alert('���ܡG�A�@��B�]�S�����A�A���������H�H�H');</script>"
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"& Session("icets")*2000 &" where id="&info(9)

Response.Write "<script Language=Javascript>alert('�z���B�G"& session("icets") &"��,��F"& session("icets")*2000 &"��ջȡI');</script>"
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
