<%'���e
function meirong()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ��O,����,¾�~,�y�O,�y�O�[,�|������,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
if left(rs("¾�~"),3)<>"���D�h" and left(rs("¾�~"),3)<>"���k�v" and left(rs("¾�~"),3)<>"���u�H" and left(rs("¾�~"),3)<>"���Ѯv" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A��¾�~�٤��O���D�h�B���k�v�B���u�H�B���Ѯv�I');}</script>"
Response.End
end if 
if rs("��O")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A����O�p��1000�A������e�I');}</script>"
Response.End
end if
if rs("����")<15 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�����Ťp��15�A�����e�I');}</script>"
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
ml=rs("�y�O")
mlj=(rs("����")*1120+2100)+rs("�y�O�[")
if rs("�y�O")>=mlj then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�{�b�n�n���A���ݭn���e�ڡI');}</script>"
Response.End
end if
if hy=1 then 
 butl=mlj-ml
 blsq=10000
 else
 butl=mlj-ml
 blsq=30000
end if

conn.execute "update �Τ� set �y�O=�y�O+"&butl&",��O=��O-"&blsq&",�ާ@�ɶ�=now() where id="&info(9)
if rs("�y�O")>mlj then
conn.execute "update �Τ� set �y�O=" & mlj & " where id="&info(9)
end if
if rs("��O")<0 then
conn.execute "update �Τ� set ��O=100 where id="&info(9)
end if
meirong="���e�v"&info(0)& "�I�i���޵��ۤv�\�ˤF�@�f�A��"& info(0) &"���y�O���t��_"& butl &"�I!"&info(0)&"��O���C"&blsq&"�I"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
