<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');window.close();}</script>"
Response.End 
end if
un=info(0)
grade=info(3)
myid=info(9)
eec=info(4)
if info(4)=0 then bbc=0 else bbc=1
mg=Request.QueryString("mg")
if instr(mg,"'")<>0 then Response.end
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("ljjh_usermdb")
conn.open connstr
sql= "select �ާ@�ɶ�,�|������ from �Τ� where �m�W='"&un&"'"
Set Rs=conn.Execute(sql)
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<60 and rs("�|������")=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��A�ާ@�I�[�J�|���L������I');location.href = 'javascript:history.go(-1)';}</script>}</script>"
	Response.End
end if
sql="select ����,�׽m�k�O,�]�k from �ϥΧޯ� where �m�W='"&un&"' and �ޯ�='"&mg&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
sql= "select * from ¾�~�ޯ� where �ޯ�='"&mg&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
learn=""
else
energy=rs("�׽m�k�O")
proviso=rs("�׽m����")
basemp=rs("���Ӫk�O")
baseat=rs("�򥻧���")
especial=rs("�]�k")
eepic=rs("����")
atdeclaration=rs("�I�k����")
sql= "select id,�ާ@�ɶ� from �Τ� where �m�W='"&un&"' and �k�O>="&energy&" and "&proviso
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
provisotxt=replace(proviso,"and","�åB")
provisotxt=replace(provisotxt,"or","�Ϊ�")
provisotxt=replace(provisotxt,">=","���C��")
provisotxt=replace(provisotxt,"<=","������")
provisotxt=replace(provisotxt,">","�j��")
provisotxt=replace(provisotxt,"<","�p��")
provisotxt=replace(provisotxt,"=","��")
learn="�׽m�ɥ��ѡA�ݪk�O"&energy&"�A�ݭn���󬰡G"&provisotxt
else
conn.Execute "insert into �ϥΧޯ�(�m�W,�ޯ�,����,�׽m�k�O,���Ӫk�O,�򥻧���,�]�k,�I�k����,����,aszcc) values('"&un&"','"&mg&"',1,"&energy&","&basemp&","&baseat&",'"&especial&"','"&atdeclaration&"','"&eepic&"','"&bbc&"')"
conn.execute "update �Τ� set �k�O=�k�O-"&energy&",�ާ@�ɶ�=now() where �m�W='"&un&"'"
learn="�׽m"&mg&"��1�Ŧ��\�C�ӥΪk�O"&energy
end if 
end if
elseif rs("����")<200 then
energy=rs("�׽m�k�O")
agrade=rs("����")
sql="select id from �Τ� where �m�W='"&un&"' and �k�O>="&energy*agrade*agrade
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
learn="�׽m�ɥ��ѡA�ݪk�O"&energy*agrade*agrade
else
conn.Execute "update �Τ� set �k�O=�k�O-"&energy*agrade*agrade&" where �m�W='"&un&"'"
conn.Execute "update �ϥΧޯ� set ����=����+1,�򥻧���=�򥻧���*1 where �m�W='"&un&"' and �ޯ�='"&mg&"'"
learn="�׽m"&mg&"��"&agrade&"�Ŧ��\�C�ӥΪk�O"&energy*agrade*agrade
end if
else learn="�׽m"&mg&"�̰ܳ��F�L�k�A��׽m�C"
end if
set rs=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<title>�ײߧޯ�</title>
</head>
<meta http-equiv="cnntent-Type" cnntent=text/html; charset=big5>
<body bgcolor='#996699' background="../bk.jpg">
<div align=center> <font color="#FFFFFF" size="2">�׽m�ޯ�</font> 
<hr>
<font color="#ffffff" size="2"><%=learn%></font> <br>
<input type=button onClick="javascript:location.href='xiulian.asp'" value='��^' name="button">
</div>
</html>