<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=3 and ���~�W='�q��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���q�ۤ���3�Ӥ���ާ@�I');</script>"
	response.end
end if
kuangid=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=5 and ���~�W='�j�T��'  "
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���j�T������5�Ӥ���ާ@�I');</script>"
	response.end
end if
shaid=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=10 and ���~�W='���Y'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����Y����10�Ӥ���ާ@�I');</script>"
	response.end
end if
muid=rs("id")
conn.Execute "update ���~ set �ƶq=�ƶq-5 WHERE id="&shaid
conn.Execute "update ���~ set �ƶq=�ƶq-3 WHERE id="&kuangid
conn.Execute "update ���~ set �ƶq=�ƶq-10 WHERE id="&muid
rs.close
rs.open "select id from �׽m���~ where ���~�W='���ߴ�' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into �׽m���~(���~�W,�֦���,�ƶq,�\��,�W�[�ƭ�) values ('���ߴ�'," & info(9) &",1,'�D�w',5000)"
else
	id=rs("id")
	conn.execute "update �׽m���~ set �W�[�ƭ�=5000,�ƶq=�ƶq+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("�z�����ߴ��׽m�����I")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>
