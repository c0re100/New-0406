<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
yilao=request.form("yilao")
if yilao<>"�L" then
'����Τ�
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �ާ@�ɶ�,��O,�Ȩ� FROM �Τ� WHERE id="& info(9) &" and ��O<1000",conn
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�t�έ���A�е�["&s&"����]�A�ާ@�I');}</script>"
	Response.End
end if	

If Rs.Bof OR Rs.Eof Then
mess="�A����i��v��"
else
sm=rs("��O")
yl=rs("�Ȩ�")
select case yilao
case "��Ǹɫ~"
bl=1.5
cd=1000-sm
case "�@��v��"
bl=1
cd=1000-sm
case "�m�ϥͩR"
bl=0.5
cd=1000-sm
end select
fy=int(cd/bl)
sm=1000
if yl<fy then
fy=yl
sm=yl*bl
end if
conn.execute "update �Τ� set ��O=��O+" & sm & ",�Ȩ�=�Ȩ�-" & fy & ",�ާ@�ɶ�=now() where id="&info(9)
mess="�g�L�J��P���v���A�A��F" & fy & "��Ȥl�W�[�ͩR��" & sm & "�I"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
else
mess="�A���Ϊv��"
end if

set rs=nothing

set conn=nothing
%>

<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000">

<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="yilao.asp">��^</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>
