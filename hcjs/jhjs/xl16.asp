<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ���� FROM �Τ� WHERE id=" & info(9),conn
if rs("����")<265 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����Ť����ڳy���Z���I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=380 and ���~�W='����'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A����������380�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id1=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=380 and ���~�W='�S�ت���'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���S�ت��ݤ���380�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id2=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=40 and ���~�W='���ݪO'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����ݪO����40�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id3=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=380 and ���~�W='�¥�'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���¥ۤ���380�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id4=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=230 and ���~�W='�Ѫ��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���Ѫ�פ���230�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id5=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=300 and ���~�W='�B��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���B������300�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id6=rs("id")
conn.Execute "update ���~ set �ƶq=�ƶq-380 WHERE id="&id1
conn.Execute "update ���~ set �ƶq=�ƶq-380 WHERE id="&id2
conn.Execute "update ���~ set �ƶq=�ƶq-40 WHERE id="&id3
conn.Execute "update ���~ set �ƶq=�ƶq-380 WHERE id="&id4
conn.Execute "update ���~ set �ƶq=�ƶq-230 WHERE id="&id5
conn.Execute "update ���~ set �ƶq=�ƶq-300 WHERE id="&id6
rs.close
rs.open "select id from ���~ where ���~�W='�ܴL�C' and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,sm,�ƶq,����) values('�ܴL�C',"&info(9)&",'����',800000,0,0,3200000,0,20032016,20032016,1,265)")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("�ܴL�C�ڳy�����I")
history.back()
</script>
