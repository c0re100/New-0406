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
if rs("����")<300 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����Ť����ڳy���Z���I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=100 and ���~�W='���_��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����_�ۤ���100�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id1=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=400 and ���~�W='�¥�'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���¥ۤ���400�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id2=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=440 and ���~�W='���ݪO'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����ݪO����440�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id3=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=200 and ���~�W='���_��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����_�ۤ���200�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id4=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=300 and ���~�W='���_��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����_�ۤ���300�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id5=rs("id")
rs.close
rs.open "SELECT id FROM ���~ WHERE �֦���=" & info(9) & " and �ƶq>=400 and ���~�W='�B��'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A���B������400�Ӥ���ާ@�I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id6=rs("id")
conn.Execute "update ���~ set �ƶq=�ƶq-100 WHERE id="&id1
conn.Execute "update ���~ set �ƶq=�ƶq-400 WHERE id="&id2
conn.Execute "update ���~ set �ƶq=�ƶq-440 WHERE id="&id3
conn.Execute "update ���~ set �ƶq=�ƶq-200 WHERE id="&id4
conn.Execute "update ���~ set �ƶq=�ƶq-300 WHERE id="&id5
conn.Execute "update ���~ set �ƶq=�ƶq-400 WHERE id="&id6
rs.close
rs.open "select id from ���~ where ���~�W='�O�s�M' and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
conn.execute "insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,sm,�ƶq,����) values ('�O�s�M',"&info(9)&",'�k��',0,0,0,3400000,900000,20032017,20032017,1,300)"

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
alert("�O�s�M�ڳy�����I")
history.back()
</script>
