<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
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
	Response.Write "<script Language=Javascript>alert('���ܡG�z��¾�~�����A�ҥH�z����b�o�̶i����I�I');</script>"
	response.end
end if
if Session("minets")>0 then
	Session("minets")=Session("minets")-1
end if
if Session("minets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�@�ڤ��Y�]�S���A�A������s�H�H�H�I');</script>"
	response.end
end if
rs.close
rs.open "select id from ���~ where ���~�W='���Y' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���Y',"&info(9)&",'���~',0,0,100,2015,2015,0,0,0,"& Session("minets") &",800,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update ���~ set �Ȩ�=800,�ƶq=�ƶq+" & Session("minets") & " where id="&id
end if
rs.close
rs.open "select �ƶq from ���~ where ���~�W='���Y' and �֦���="&info(9)
Session("minesl")=rs("�ƶq")
Response.Write "<script Language=Javascript>alert('�z�����Y�G"& session("minets") &"��,�{�֦����Y"& rs("�ƶq") &"�ڡA�h�h�V�O�I');</script>"
Session("minets")=0
Session("cbsj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
