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
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ�,����,��O�[,���O�[,��O,���O,���� from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'fangniu.asp';}</script>"
	Response.End
end if
tlj=(rs("����")*1500+5260)+rs("��O�[")
nlj=(rs("����")*640+2000)+rs("���O�[")
tla=rs("��O")
nla=rs("���O")
mymoney=rs("����")
dj=rs("����")
if mymoney<200 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('���~�I�A�ثe�����C��200,�v�ũȧA������ɭԤ��p�߱��U�s�V�A�ҥH����L�H���|�F�I');location.href = 'fangniu.asp';}</script>"

response.end
end if
if dj<2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('���~�I�A�ثe�����ŧC��3,�v�ũȧA������ɭԹJ�쳥�~�A�ҥH�����A�h�F�I');location.href = 'fangniu.asp';}</script>"
response.end
end if
randomize timer
r=int(rnd*10)+1
select case r
case 1
randomize timer
s=int(rnd*50)
mess="���ߧA������A���F�@���Z�L�_�ѩM��賣�W�["& s &""
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+'"& s &"',����=����+'"& s &"',�ާ@�ɶ�=now()  where id="&info(9)
case 2
randomize timer
s=int(rnd*50000)
	mess="���ߧA������A�b����ߨ�"& s &"��"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+'"& s &"',�ާ@�ɶ�=now()  where id="&info(9)
case 3
	mess="�A������M�ߡA������F�A�v�Ż@�A5000��A�y�O���100�I"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-5000,�y�O=�y�O-100,�ާ@�ɶ�=now()  where id="&info(9)
case 4
	mess="���ߧA������A�b����ߨ�@�Ө��l�A�I�ƼW�[2�I"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-5000,�y�O=�y�O-100,�w���I��=�w���I��+2,�ާ@�ɶ�=now()  where id="&info(9)
case 5
randomize timer
s=int(rnd*5000)
	mess="���߱z!�g�L�@�Ѫ�����A�v�ŵ��F�A�@���Z�ǯ��D�A�ϧA��o"& s &"��O,���O����"& s &"�A���W�S���նO�����I�I"
if tla>=tlj or nla>=nlj then
conn.execute "update �Τ� set ��O='"&tlj&"',���O='"& nlj &"',�ާ@�ɶ�=now()  where id="&info(9)
else
conn.execute "update �Τ� set ��O=��O+'"& s &"',���O=���O+'"& s &"',�ާ@�ɶ�=now()  where id="&info(9)
end if
case 6
rs.close
rs.open "select id from ���~ where ���~�W='����' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,sm,�|��) values ('����','"&info(9)&"','�r��',0,0,0,0,10000,1,20000,100013,100013,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
 end if
mess="���ߧA������A�b����ߨ�@�Ӥ���,�i���˹����O10000"
case 7
	mess="����"&info(0)&"����ɶ^��s�}�̡A���t���X�Q�A�o�{�@���U�~�F�ۡA�A�Υi�W�[��O200000�I"
rs.close
rs.open "select id from ���~ where ���~�W='�U�~�F��' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,����,���s,��T��,����,sm,�ƶq,�Ȩ�,�|��) values ('�U�~�F��',"&info(9)&",'�ī~',0,200000,0,0,100,136,136,1,5000000,"&aaao&")"
                        else 
id=rs("id")
				sql="update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�U�~�F��' and id="&id
				conn.execute(sql)
		        end if
case 8
	mess="����"&info(0)&"����ɦb����ߨ�@���m�ư��x�Сn"
rs.close
rs.open "select id from ���~ where ���~�W='�ư��x��' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,����,���s,��T��,����,sm,�ƶq,�Ȩ�,�|��) values ('�ư��x��',"&info(9)&",'���~',0,0,0,0,100,99002,99002,1,5000000,"&aaao&")"
                        else 
id=rs("id")
				sql="update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�ư��x��' and id="&id
				conn.execute(sql)
		        end if
case 9
	mess="����"&info(0)&"����ɦb����ߨ�@���X���A�����i�Ψӷڳy���@�Z��!"
rs.close
rs.open "select id from ���~ where ���~�W='�X����' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,����,���s,��T��,����,sm,�ƶq,�Ȩ�,�|��) values ('�X����',"&info(9)&",'���~',0,0,0,0,0,2003302,2003302,1,10000,"&aaao&")"
                        else 
id=rs("id")
				sql="update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�X����' and id="&id
				conn.execute(sql)
		        end if
case else
rs.close
rs.open "select id from ���~ where ���~�W='����' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,sm,�|��) values ('����','"&info(9)&"','�r��',0,0,0,0,10000,1,20000,100013,100013,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
 end if
mess="���ߧA������A�b����ߨ�@�Ӥ���,�i���˹����O10000"
end select
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'fangniu.asp';}</script>"

response.end%>