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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select count(*) from �Τ� where ���ФH='"&info(0)&"'",conn
jsren=rs(0)
rs.close
rs.open "Select * from ���X where id=1",conn
gfmoney=rs("gfmoney")
zmmoney=rs("zmmoney")
zlmoney=rs("zlmoney")
tzmoney=rs("tzmoney")
dzmoney=rs("dzmoney")
lznum=rs("lznum")
zsnum=rs("zsnum")
rs.close
set rs=conn.execute ("Select count(*) from �Τ� where ����='"&info(5)&"'")
renshu=rs(0)-1
rs.close
rs.open "Select ���,�|������,����,����,�Ȩ�,�s��,���� from �Τ� where id="&info(9),conn
mdate=rs("���")
hydd=rs("�|������")
shenfen=rs("����")
gf=rs("����")
if  DateDiff("d",rs("���"),now())=0 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ѧr�A["&info(0) &"]���ѧA�ӻ�L�����I�ѰO�F�H');window.close();</script>"
response.end
end if
yin=rs("�Ȩ�")+rs("�s��")
if yin>(rs("�|������")+1)*500000000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ѧr�A["&info(0) &"]�A�{�b���o��h���Ȩ�A�ٷQ�n�r�I�I');window.close();</script>"
response.end
end if
dj=rs("����")
hy=rs("�|������")
if trim(shenfen)="�x��" then

	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+zmmoney
	conn.execute("update �Τ� set �Ȩ�=�Ȩ�+"&money&",���=now() where id="&info(9))
	
end if
if trim(shenfen)="����" then
	'�����
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+zlmoney
	conn.execute("update �Τ� set �Ȩ�=�Ȩ�+"&money&",���=now() where id="&info(9))
	
end if
if trim(shenfen)="��D" then
	'�����
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+tzmoney
	conn.execute("update �Τ� set �Ȩ�=�Ȩ�+"&money&",���=now() where id="&info(9))
	
end if
if trim(gf)="�x��" then
    '�����
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+gfmoney
	conn.execute("update �Τ� set �Ȩ�=�Ȩ�+"&money&",���=now() where id="&info(9))
end if
if trim(shenfen)<>"�x��" and trim(shenfen)<>"����" and trim(shenfen)<>"��D" and trim(gf)<>"�x��" then
	'�����
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+dzmoney
	conn.execute("update �Τ� set �Ȩ�=�Ȩ�+"&money&",���=now() where id="&info(9))
end if
rs.close
rs.open "select id from ���~ where ���~�W='�U�~�F��' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�U�~�F��',"&info(9)&",'�ī~',0,200000,2,5000000,136,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update ���~ set �ƶq=�ƶq+"&lznum&",�|��="&aaao&" where id="& id &" and �֦���="&info(9)
				
		        end if
if hy>0 then
rs.close
	rs.open "select id from ���~ where ���~�W='�]�O�p��' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�]�O�p��',"&info(9)&",'�k��',0,0,2,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+"&zsnum&",�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
end if
rs.close
rs.open "Select ����,����,���� from �Τ� where id="&info(9),conn
%>
<html>
<head>
<title>�u�����B</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE type=text/css>
TD {FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "�s�ө���"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#0099FF" text="#FFFFFF" leftmargin="0">
<div align="center">
<p><font color="#000000"><b><font size="+3">�u�����B</font></b></font> <br>
<br>
<font size=+1> <b><%=info(0)%></b></font>�A�O<font color="#0000FF">[<%=rs("����")%>]</font>
��<font color="#FF0000"> [<%=rs("����")%>] </font>�ԤH�G<font color="#FF0000"><b><%=jsren%>��</b></font>
�԰�����:<font color="#FF0000"><%=rs("����")%>��</font>
<%if trim(rs("����"))="�x��" then%>
�A���̤l:<font color="#FF0000"><b><%=renshu%>��</b></font> <%end if%><br>�h�h�V�O�A�h�h�ԤH�I
<br>
���ѧA�q��������F<b><font color="#FF0000"><%=money%>��</font></b>�A�p�߫O�s�A���n�ê�I
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
<br>
</p>
  <p align="center"><br>
    <font color="#FF0000" size="+1"><b>�u��p���k</b></font><br>
    <font color="#0000FF">�x���G</font>�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500+<%=gfmoney%> 
    <br>
    <font color="#0000FF">�x���G</font>�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500+<%=zmmoney%><br>
    <font color="#0000FF">���ѡG</font>�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500+<%=zlmoney%><br>
    <font color="#0000FF">��D�G</font>�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500+<%=tzmoney%> 
    <br>
    <font color="#0000FF">�̤l�G</font>�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500+<%=dzmoney%> 
    <br>
    <font color="#0000FF">�|���G</font>�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500+<%=dzmoney%> 
    <br>
    <br>
    <font color="#99FFCC">�D�|���C�ѥi�H�b�����<%=lznum%>�ӸU�~�F��<br>
    �|���C�ѥi�H�b�����<%=zsnum%>���]�O�p��</font></p> </p>
  <p align="center">&nbsp;</p>
<p align="center">
<input type=button value=�������f onClick='window.close()' name="button">
</p>
</div>
</body>
</html>
