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
if info(4)=0 then 
aaao=0
else
aaao=1
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
	Response.Write "<script Language=Javascript>alert('���ܡG�z��¾�~����b�o�̶i����B�A���ഫ¾�~�I');</script>"
	response.end
end if
if Session("icets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�@��B�]�S�����A�s����H');</script>"
	response.end
end if
randomize timer
r=int(rnd*3)
if r=2 then
rs.close
rs.open "select id from ���~ where ���~�W='����' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('����',"&info(9)&",'���~',0,0,0,2003306,2003306,0,0,0,1,10000,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if
Response.Write "<script Language=Javascript>alert('�z����@�����ۡA�����i�H�Ψӷڳy���@�Z���I');</script>"
end if
rs.close
rs.open "select id from ���~ where ���~�W='�B��' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�B��',"&info(9)&",'�ī~',0,0,100,151,151,"& Session("icets") &",2000,151,151,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update ���~ set �Ȩ�=2000,�ƶq=�ƶq+" & Session("icets") & ",�|��="&aaao&" where id="&id
end if
rs.close
rs.open "select �ƶq from ���~ where ���~�W='�B��' and �֦���="&info(9),conn
Session("icesl")=rs("�ƶq")
Response.Write "<script Language=Javascript>alert('�z���B�G"& session("icets") &"��,�{�֦��B��"& rs("�ƶq") &"��A�h�h�V�O�I');</script>"
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
