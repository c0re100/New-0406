<%'��\�v
function qigongshi()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ����,¾�~,���O,��O,���O�[,��O�[,�|������,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
if left(rs("¾�~"),3)<>"���k�v" and left(rs("¾�~"),3)<>"���u�H" and left(rs("¾�~"),3)<>"���Ѯv" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A��¾�~�٤��O���k�v�B���u�H�B���Ѯv���ɮ�I');}</script>"
Response.End
end if  
if rs("��O")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A����O�p��1000�A�L�k�v���I');}</script>"
Response.End
end if
if rs("����")<40 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�����Ťp��40�A�L�k�v���I');}</script>"
Response.End
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
if rs("�|������")>0 then 
 hy=1
else
 hy=0
 end if
nl=rs("���O")
tlj=(rs("����")*1500+5260)+rs("��O�[")
nlj=(rs("����")*640+2000)+rs("���O�[")
if rs("���O")>=nlj then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�{�b����n�n���A���ݭn�A���v���ڡI');}</script>"
Response.End
end if
if hy=1 then 
 butl=nlj-nl
 blsq=1000
 else
 butl=nlj-nl
 blsq=5000
end if
conn.execute "update �Τ� set ���O=���O+"&butl&",��O=��O-"&blsq&",�ާ@�ɶ�=now() where id="&info(9)
if rs("���O")>nlj then
conn.execute "update �Τ� set ���O=" & nlj & " where id="&info(9)
end if
if rs("��O")<0 then
conn.execute "update �Τ� set ��O=100 where id="&info(9)
end if
qigongshi="��\�j�v"&info(0)& "��ۤv�ϥΤF�v���N�A"&info(0)&"�����O���t��_"& butl &"�I!"&info(0)&"�l��"&blsq&"����O�I"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
