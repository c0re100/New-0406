<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../../images/7.jpg">
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
sl=int(abs(Request.form("sl")))
if sl<1 or sl>9 then
	Response.Redirect "../../error.asp?id=72"
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ����,���~�W,�Ȩ�,���O,��O,����,���s,��T��,sm,���� FROM ���~�R where ID=" & ID,conn
if rs("����")<>"�r��" then
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
if info(4)>1 then
	yin=int(yin/2)
end if
rs.close
rs.open "select �Ȩ�,�ާ@�ɶ� from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if yin*sl>rs("�Ȩ�") then
	Response.Write "<script Language=Javascript>alert('���ܡG�ʶR�����\�A��]�G�A���Ȩ⤣���F');location.href = 'javascript:history.go(-1)';</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin*sl & ",�ާ@�ɶ�=now() where id="&info(9)
rs.close
'���~
rs.open "select id from ���~ where ���~�W='" & wu & "' and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
	sql="insert into ���~(���~�W,�֦���,����,���O,��O,����,���s,��T��,�ƶq,�Ȩ�,����,sm,����,�|��) values ('"&wu&"',"&info(9)&",'"&lx&"',"&nl&","&tl&","&gj&","&fy&","&jgd&","&sl&","&int(yin*2)&",'�L',"&sm&","&dj&","&aaao&")"
	conn.execute sql
else
	id=rs("id")
	sql="update ���~ set �ƶq=�ƶq+" & sl & ",�|��="&aaao&" where id="&id
	conn.execute sql
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
if info(4)>1 then
	Response.Write "<script Language=Javascript>alert('���ܡG2�ťH�W�|���ʶR���~��5��,�ʶR"&wu&sl&"�Ӧ��\�I');location.href = 'javascript:history.go(-1)';</script>"
response.end
else
	Response.Write "<script Language=Javascript>alert('���ܡG�ʶR"&wu&"���~"&sl&"�Ӧ��\,�p�G�z�O2�ťH�W�|���i�H��5��I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
%>
