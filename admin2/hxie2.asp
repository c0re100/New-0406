<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�|���O,���O,��O,����,���s,�|������,�ާ@�ɶ�,���O�[,��O�[,�����[,���s�[ from �Τ� where id="&info(9),conn
nlj=(rs("����")*640+2000)+rs("���O�[")
nla=rs("���O")
tlj=(rs("����")*1500+5260)+rs("��O�[")
tla=rs("��O")
gjj=(rs("����")*1200+4500)+rs("�����[")
gja=rs("����")
fyj=(rs("����")*1120+3000)+Rs("���s�[")
fya=rs("���s")
hy=rs("�|������")
hf=rs("�|���O")
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'hxie.asp';}</script>"
	Response.End
end if
if hy=0 then
Response.Write "<script language=JavaScript>{alert('�A���O�|���Ц^�a�I');location.href = '../help/info.asp?type=2&name=�|���[�J��k';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
else
cid=Request.QueryString("id")
select case cid
Case "1"
if nla>=nlj then
conn.execute("update �Τ� set ��O=��O-1000,���O="& nlj &",�ާ@�ɶ�=now()  where id="&info(9))
else
conn.execute("update �Τ� set ���O=���O+1000,��O=��O-1000,�ާ@�ɶ�=now()  where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�A��1000����O���F1000�����O');location.href='hxie.asp';</script>"
Response.End
Case "2"
if tla>=tlj then
conn.execute("update �Τ� set ���O=���O-1000,��O="& tlj &",�ާ@�ɶ�=now()  where id="&info(9))
else
conn.execute("update �Τ� set ���O=���O-1000,��O=��O+1000,�ާ@�ɶ�=now()  where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�A��1000�����O���F1000����O');location.href='hxie.asp';</script>"
Response.End
Case "3"
if fya>=fyj then
conn.execute("update �Τ� set ����=����-1000,���s="& fyj &",�ާ@�ɶ�=now()  where id="&info(9))
else
conn.execute("update �Τ� set ����=����-1000,���s=���s+1000,�ާ@�ɶ�=now()  where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�A��1000���������F1000�����s');location.href='hxie.asp';</script>"
Response.End
Case "4"
if gja>=gjj then
conn.execute("update �Τ� set ���s=���s-1000,����="& gjj &",�ާ@�ɶ�=now()  where id="&info(9))
else
conn.execute("update �Τ� set ���s=���s-1000,����=����+1000,�ާ@�ɶ�=now()   where id="&info(9))
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�A��1000�����s���F1000������');location.href='hxie.asp';</script>"
Response.End
Case "5"
if rs("�|���O")>=1 then
jb=int(abs(hf)*100)
conn.execute("update �Τ� set �|���O=0,����="&jb&",�ާ@�ɶ�=now()  where id="&info(9))
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�|�O�ഫ���������\�A�Ъ`�N�d��!');location.href='hxie.asp';</script>"
Response.End
else

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�A�٦��|�O�ܡH�H�H');location.href='hxie.asp';</script>"
Response.End
end if
Case else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('���~����');location.href='hxie.asp';</script>"
Response.End
end select
end if%>
