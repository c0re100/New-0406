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
rs.open "SELECT id  FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=10 and ���~�W='�q��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���q�ۤ���10�Ӥ���ާ@�I');</script>"
	response.end
end if
kuangid=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=2 and ���~�W='�p�U��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���p�U������2�Ӥ���ާ@�I');</script>"
	response.end
end if
liyuid=rs("id")
conn.Execute "update ���~ set �ƶq=�ƶq-10 WHERE id="&kuangid
conn.Execute "update ���~ set �ƶq=�ƶq-2 WHERE id="&liyuid
rs.close
rs.open "select id from �׽m���~ where ���~�W='�ܻT�P��' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into �׽m���~(���~�W,�֦���,�ƶq,�\��,�W�[�ƭ�) values ('�ܻT�P��',"& info(9) &",1,'�y�O',4000)"
else
	id=rs("id")
	conn.execute "update �׽m���~ set �W�[�ƭ�=4000,�ƶq=�ƶq+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("�z���ܻT�P���׽m�����I")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>