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
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"work/famu")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ¾�~ from �Τ� where id="&info(9),conn
if trim(rs("¾�~"))="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�z��¾�~�����A�ҥH�z����b�o�̬����I�I');</script>"
	response.end
end if
if Session("minets")=1 then
	Session("minets")=Session("minets")-1
end if
if Session("minets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����Y�Ӥ֡A�椣�F�X�ӿ��H�H�H�I');</script>"
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"& Session("minets")*800 &" where id="&info(9)
Response.Write "<script Language=Javascript>alert('�z�����Y�G"& session("minets") &"��,��F"& session("minets")*800 &"��ջȡI');</script>"
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
