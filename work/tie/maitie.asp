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
rs.open "select ¾�~ from �Τ� where id="&info(9),conn

if Session("tiets")<=1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�@���K�]�S���A�A���������H�H�H�I');</script>"
	response.end
end if
if Session("tiesl")<=0 then
	Session("tiesl")=0
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"& Session("tiets")*3000 &" where id="&info(9)
conn.execute "update ���~ set �ƶq="& Session("tiesl") &" where ���~�W='�q��' and �֦���="&info(9)
Response.Write "<script Language=Javascript>alert('�z���K���G"& session("tiets") &"��,��F"& session("tiets")*3000 &"��ջȡI');</script>"
Session("tiets")=0
Session("tiesj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>