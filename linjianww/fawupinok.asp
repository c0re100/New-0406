<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=trim(request.form("a"))
b=trim(request.form("b"))
c=trim(request.form("c"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM �Τ� where �m�W='"&a&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�S���ӥΤ�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
idd=rs("id")
select case b

case "���C"
rs.close
rs.open "select id from ���~ where ���~�W='���C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('���C',"&idd&",'�k��',0,0,0,100000,100000,2003201,"&c&",5,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�����C"
rs.close
rs.open "select id from ���~ where ���~�W='�����C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�����C',"&idd&",'����',100000,0,0,300000,0,2003202,"&c&",15,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�Q�ڼC"
rs.close
rs.open "select id from ���~ where ���~�W='�Q�ڼC' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�Q�ڼC',"&idd&",'�k��',0,0,0,500000,200000,2003203,"&c&",25,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "���Z�C"
rs.close
rs.open "select id from ���~ where ���~�W='���Z�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('���Z�C',"&idd&",'����',200000,0,0,800000,0,2003204,"&c&",35,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�ȭ߼C"
rs.close
rs.open "select id from ���~ where ���~�W='�ȭ߼C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�ȭ߼C',"&idd&",'�k��',0,0,0,1000000,300000,2003205,"&c&",45,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�H�P�]�C"
rs.close
rs.open "select id from ���~ where ���~�W='�H�P�]�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�H�P�]�C',"&idd&",'����',300000,0,0,1200000,0,2003206,"&c&",55,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�涳���b"
rs.close
rs.open "select id from ���~ where ���~�W='�涳���b' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�涳���b',"&idd&",'�k��',0,0,0,1400000,400000,2003207,"&c&",65,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�ݦ�C"
rs.close
rs.open "select id from ���~ where ���~�W='�ݦ�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�ݦ�C',"&idd&",'����',400000,0,0,1600000,0,2003208,"&c&",75,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�C�Y�]�C"
rs.close
rs.open "select id from ���~ where ���~�W='�C�Y�]�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�C�Y�]�C',"&idd&",'�k��',0,0,0,1800000,500000,2003209,"&c&",85,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�ĳ��C"
rs.close
rs.open "select id from ���~ where ���~�W='�ĳ��C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�ĳ��C',"&idd&",'����',500000,0,0,2000000,0,20032010,"&c&",95,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "���s�C"
rs.close
rs.open "select id from ���~ where ���~�W='���s�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('���s�C',"&idd&",'�k��',0,0,0,2200000,600000,20032011,"&c&",105,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�C�P�_�C"
rs.close
rs.open "select id from ���~ where ���~�W='�C�P�_�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�C�P�_�C',"&idd&",'����',600000,0,0,2400000,0,20032012,"&c&",125,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�������M"
rs.close
rs.open "select id from ���~ where ���~�W='�������M' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�������M',"&idd&",'�k��',0,0,0,2600000,700000,20032013,"&c&",155,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�񲴼C"
rs.close
rs.open "select id from ���~ where ���~�W='�񲴼C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�񲴼C',"&idd&",'����',700000,0,0,2800000,0,20032014,"&c&",205,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "���ѼC"
rs.close
rs.open "select id from ���~ where ���~�W='���ѼC' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('���ѼC',"&idd&",'�k��',0,0,0,3000000,800000,20032015,"&c&",245,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�ܴL�C"
rs.close
rs.open "select id from ���~ where ���~�W='�ܴL�C' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values('�ܴL�C',"&idd&",'����',800000,0,0,3200000,0,20032016,"&c&",265,"&aaao&")")

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�O�s�M"
rs.close
rs.open "select id from ���~ where ���~�W='�O�s�M' and �֦���="&idd,conn
If Rs.Bof OR Rs.Eof then
conn.execute "insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,�|��) values ('�O�s�M',"&idd&",'�k��',0,0,0,3400000,900000,20032017,"&c&",300,"&aaao&")"

else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="&id
end if
case "�s�]"
rs.close
rs.open "select id from ���~ where ���~�W='�s�]' and �֦���=" & idd,conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,sm,�|��) values ('�s�]',"&idd&",'���~',0,100,0,0,0,"&c&",10000000,50000,0,50000,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+"&c&",�|��="&aaao&" where id="& id &" and �֦���="&info(9)
				
		        end if
end select
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�o���~�ާ@')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
