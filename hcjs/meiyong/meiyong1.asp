
<body bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ�,�y�O,���A from �Τ� where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�A�A���O���򤤤H');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
else
mymoney=rs("�Ȩ�")
if rs("�y�O")<1 or rs("�Ȩ�")>100 and rs("���A")<>"��" and rs("���A")<>"��" then
randomize timer
r=int(rnd*3)
if r=1 then
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-100,�y�O=�y�O+12 where id="&info(9)
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "�g�L���e�|�p�j���@�f�\��,�A���y�O�W�[�F12�I,���e�|�p�j�����F�A100��Ȥl!"
response.end
else
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-10 where id="&info(9)
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "���e�|���Ǯ{�ѩ��������,�ˤF�b�Ѥ]�S�����A�W�[�y�O.�ժ�F�A�Q��A�ȶO!"
response.end
end if
else
response.write "�S���]�Q���?���e�|�p�j�̤@�f�տا�A�F�����f"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
%>
<%end if%>
