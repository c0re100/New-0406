<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<body topmargin="6" leftmargin="0" background="../../bg.gif">
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,���~�W,�Ȩ�,����,���O,��O,����,���s,��T��,sm,���� FROM ���~�R where ID=" & ID,conn
if rs("����")<>"�A��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=72"
	response.end
end if
wu=rs("���~�W")
yin=rs("�Ȩ�")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
sm=rs("sm")
gj=rs("����")
fy=rs("���s")
jgd=rs("��T��")
dj=rs("����")
if info(4)>0 then
	yin=int(yin/2)
end if
rs.close
rs.open "select �Ȩ�,�ާ@�ɶ� from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if yin>rs("�Ȩ�") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�ʶR�����\�A��]�G�A���Ȩ⤣���F');window.close();</script>"
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & ",�ާ@�ɶ�=now() where id="&info(9)
rs.close
'���~
rs.open "select id from ���~ where ���~�W='" & wu & "' and �֦���="& info(9),conn
If Rs.Bof OR Rs.Eof then
	sql="insert into ���~(���~�W,�֦���,����,���O,��O,����,���s,��T��,�ƶq,�Ȩ�,����,sm,����,�|��) values ('"&wu&"',"&info(9)&",'"&lx&"',"&nl&","&tl&","&gj&","&fy&","&jgd&",1,"&int(yin*2)&",'�L',"&sm&","&dj&","&aaao&")"
	conn.execute sql
else
	id=rs("id")
	sql="update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id
	conn.execute sql
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if info(4)>0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�|���ʶR���~��5��,�ʶR["&wu&"]1�Ӧ��\�I');window.close();</script>"
response.end
else
Response.Write "<script Language=Javascript>alert('���ܡG�ʶR["&wu&"]���~1�Ӧ��\,�p�G�z�O�|���i�H��5��I');window.close();</script>"
response.end
end if
%> 