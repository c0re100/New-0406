<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../../images/7.jpg">
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
sl=int(abs(Request.form("sl")))
if sl<1 or sl>9 then
	Response.Redirect "../../error.asp?id=72"
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
select case id

case 1
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if yin*sl<10000000 or jin*sl<1000 then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�l�u' and �֦���=" &info(9),conn
If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�l�u',"&info(9)&",'�u��',0,0,100,2012,2012,0,0,0,"&sl&",2000000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
	end if
case 2
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*20000000>yin or sl*5000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�Ѫ��' and �֦���=" &info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('�Ѫ��',"&info(9)&",'�ī~',0,0,0,150,150,"&sl&",2000,400,400,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update ���~ set �Ȩ�=2000,�ƶq=�ƶq+1,�|��="&aaao&" where  id="&id
end if
case 3
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*30000000>yin or sl*10000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�T����' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�T����',"&info(9)&",'�k��',0,0,1000,2001,2001,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id 
	end if
case 4
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*40000000>yin or sl*20000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�}���@' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�}���@',"&info(9)&",'�k��',0,0,1000,2002,2002,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 5
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*50000000>yin or sl*30000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='��w�l' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('��w�l',"&info(9)&",'�k��',0,0,1000,2004,2004,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 6
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*60000000>yin or sl*40000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�m�T�O' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�m�T�O',"&info(9)&",'�k��',0,0,1000,2005,2005,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 7
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*70000000>yin or sl*50000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='���_��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���_��',"&info(9)&",'�k��',0,0,1000,2006,2006,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 8
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*80000000>yin or sl*60000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='���_��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���_��',"&info(9)&",'�k��',0,0,1000,2007,2007,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 9
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*90000000>yin or sl*70000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='���_��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���_��',"&info(9)&",'�k��',0,0,1000,2008,2008,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 10
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*100000000>yin or sl*80000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�]�O�p��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�]�O�p��',"&info(9)&",'�k��',0,0,100,1100,1100,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if

case 11
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*200000000>yin or sl*90000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�ͤ�J�|' and �֦���=" &info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�ͤ�J�|',"&info(9)&",'�k��',0,0,100,2009,2009,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id 
	end if
case 12
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*300000000>yin or sl*100000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�����M' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�����M',"&info(9)&",'�k��',0,0,1000,2010,2010,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 13
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*400000000>yin or sl*110000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�����y' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('�����y',"&info(9)&",'�k�_',100,0,0,0,0,"&sl&",200000,1600,1600,0,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 14
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*500000000>yin or sl*120000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�j�T��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j�T��',"&info(9)&",'�t��',0,0,100,300,300,0,-1200,-800,"&sl&",20000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=20000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 15
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*600000000>yin or sl*130000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�j��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j��',"&info(9)&",'�ī~',0,0,100,301,301,0,300,50,"&sl&",2000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=2000,�ƶq=�ƶq+1,�|��="&aaao&" where id="&id
	end if
case 16
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*700000000>yin or sl*140000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�p�U��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�p�U��',"&info(9)&",'�ī~',0,0,100,302,302,0,100,30,"&sl&",500,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=500,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id
	end if
case 17
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*800000000>yin or sl*150000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�ƶ����y' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('�ƶ����y',"&info(9)&",'���~',0,0,0,0,0,"&sl&",200000,99004,99004,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id

	end if
case 18
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*900000000>yin or sl*160000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�S�ت���' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('�S�ت���',"&info(9)&",'���~',0,0,0,0,0,"&sl&",10000,2003303,2003303,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id

	end if
case 19
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*100000000>yin or sl*170000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='���믵�Z' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('���믵�Z',"&info(9)&",'���~',0,0,0,0,0,"&sl&",200000,99003,99003,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id

	end if
case 20
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1100000000>yin or sl*180000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�B�v�C��' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('�B�v�C��',"&info(9)&",'���~',0,0,0,0,0,"&sl&",200000,99005,99005,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id

	end if
case 21
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1200000000>yin or sl*190000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�p�^���y' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
  conn.execute "insert into ���~(���~�W,�֦���,����,���s,��O,���O,��T��,����,����,sm,�ƶq,�Ȩ�,����,�|��)  values('�p�^���y',"&info(9)&",'���~',0,0,0,40000,15000,99006,99006,"&sl&",10000000,0,"&aaao&")" 
else
id=rs("id")
conn.execute  "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&"  where id="& id
end if
case 22
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1300000000>yin or sl*200000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�¥�' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
  conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��O,���O,��T��,����,sm,�ƶq,�Ȩ�,����,�|��) values('�¥�',"&info(9)&",'���~',0,0,0,0,0,2003305,2003305,"&sl&",10000,0,"&aaao&")" 
else
id=rs("id")
 conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id 
end if

case 23
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1400000000>yin or sl*210000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='���ݪO' and �֦���=" & info(9),conn
if rs.eof and rs.bof then
  conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��O,���O,��T��,����,sm,�ƶq,�Ȩ�,����,�|��)  values('���ݪO',"&info(9)&",'���~',0,0,0,0,0,2003304,2003304,"&sl&",10000,0,"&aaao&")" 
else
id=rs("id")
 conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id 
end if

case 24
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1500000000>yin or sl*220000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='����' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('����',"&info(9)&",'���~',0,0,0,2003306,2003306,0,0,0,"&sl&",10000,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if

case 25
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1600000000>yin or sl*230000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�B��' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��)  values ('�B��',"&info(9)&",'�ī~',0,0,100,151,151,"&sl&",2000,151,151,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if

case 26
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1700000000>yin or sl*240000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�K��' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�K��',"&info(9)&",'���~',0,0,0,2003301,2003301,0,0,0,"&sl&",10000,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if

case 27
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1800000000>yin or sl*250000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='�q��' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�q��',"&info(9)&",'���~',0,0,0,2014,2014,0,0,0,"&sl&",800,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if

case 28
rs.open "select �Ȩ�,���� from �Τ� where id=" &info(9),conn
yin=rs("�Ȩ�")
jin=rs("����")
if sl*1900000000>yin or sl*260000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�S������h���M�����I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",����=����-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from ���~ where ���~�W='���Y' and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���Y',"&info(9)&",'���~',0,0,100,2015,2015,0,0,0,"&sl&",800,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
end if
case else
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ާ@���~�I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ʶR���~���\�I');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>
