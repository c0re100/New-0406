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
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=8 and ���~�W='�B��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���B������8�Ӥ���ާ@�I');</script>"
	response.end
end if
iceid=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=8 and ���~�W='�q��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���q�ۤ���8�Ӥ���ާ@�I');</script>"
	response.end
end if
kuangid=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=3 and ���~�W='�j�T��'  "
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���j�T������3�Ӥ���ާ@�I');</script>"
	response.end
end if
sayuid=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=2 and ���~�W='�Ѫ��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���Ѫ�פ���2�Ӥ���ާ@�I');</script>"
	response.end
end if
laohuid=rs("id")
conn.Execute "update ���~ set �ƶq=�ƶq-8 WHERE id="&iceid
conn.Execute "update ���~ set �ƶq=�ƶq-8 WHERE id="&kuangid
conn.Execute "update ���~ set �ƶq=�ƶq-3 WHERE id="&sayuid
conn.Execute "update ���~ set �ƶq=�ƶq-2 WHERE id="&laohuid
rs.close
rs.open "select id from �׽m���~ where ���~�W='�z�L�����Y' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into �׽m���~(���~�W,�֦���,�ƶq,�\��,�W�[�ƭ�) values ('�z�L�����Y'," & info(9) &",1,'�Z�\',15000)"
else
	id=rs("id")
	conn.execute "update �׽m���~ set �W�[�ƭ�=15000,�ƶq=�ƶq+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("�z�z�L�����Y�׽m�����I")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>