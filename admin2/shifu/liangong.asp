<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select �v��,�Z�\�[,��O,����,����,�ާ@�ɶ� from �Τ� where ID=" & info(9)
set rs=conn.execute(sql)
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'shifu.asp';}</script>"
	Response.End
end if
wushu=rs("�Z�\�[")
tili=rs("��O")
zhizi=rs("����")
sf=rs("�v��")
if sf="�L"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�٨S���v�ŰڡA�h���v�a�I');location.href = 'shifu.asp';}</script>"
Response.End
end if
if tili<110 or zhizi<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
  MsgBox "�A�����O�p��100�Ϊ̸��p��20�A�N�Q�m�W�����Z�\�A����򥻥\�m�n�a"
  location.href = "shifu.asp"
</script><%
else
if rs("�Z�\�[")<=rs("����")*1000 then
	conn.execute "update �Τ� set ��O=��O-100,����=����-20,�Z�\�[=�Z�\�[+10,�ާ@�ɶ�=now() where id="&info(9)	
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�{�b�A���W�����F�A���ɤF�ŦA�m�a');location.href = 'shifu.asp';}</script>"
	Response.End
end if
'conn.execute "update �Τ� set ��O=��O-100,����=����-20,�Z�\�[=�Z�\�[+10 where id="&info(9)
%>
<script language="vbscript">
  MsgBox "�g�L�v�Ū��ձСA�A�ᱼ���20�I�A�A���m�Z�����W��10�I�C"
  location.href = "shifu.asp"
</script><%
end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing%>

