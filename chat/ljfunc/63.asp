<%'�x�x
function junguan()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ����,¾�~,�D�w,�Z�\,�Z�\�[,�|������,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
if left(rs("¾�~"),3)<>"���u�H" and left(rs("¾�~"),3)<>"���Ѯv" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A��¾�~�٤��O���u�H�B���Ѯv����_�Z�\�I');}</script>"
Response.End
end if 
if rs("�D�w")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���D�w�p��1000�A�i�J���F�����}�I');}</script>"
Response.End
end if
if rs("����")<40 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�����Ťp��40�A�i�J���F�����}�I');}</script>"
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
wgj=(rs("����")*1280+3800)+rs("�Z�\�[")
wg=rs("�Z�\")
if rs("�Z�\")>=wgj then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�Z�\�ܰ��F�ڡA���ݭn�A���ɤF�ڡI');}</script>"
Response.End
end if
if hy=1 then 
 butl=wgj-wg
 blsq=10000
 else
 butl=wgj-wg
 blsq=50000
end if

conn.execute "update �Τ� set �Z�\=�Z�\+"&butl&",�D�w=�D�w-"&blsq&",�ާ@�ɶ�=now() where id="&info(9)
if rs("�Z�\")>wgj then
conn.execute "update �Τ� set �Z�\=" & wgj & " where id="&info(9)
end if
if rs("�D�w")<0 then
conn.execute "update �Τ� set �D�w=100 where id="&info(9)
end if

junguan="�x�x"&info(0)& "���F���ɦۤv���Z�\�A�����_�ۨ����J�]���M�I�i�J�����}�A�ש󦨥\��_�Z�\"& butl &"�I!"&from1&"�D�w���C"&blsq&"�I"
rs.close
conn.close
set rs=nothing	
set conn=nothing
end function
%>
