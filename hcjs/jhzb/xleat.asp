<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhzb")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
sl=abs(int(Request.form("xlsl")))
if sl<1 or sl>50 then
	Response.Redirect "../../error.asp?id=72"
end if
action=request.querystring("action")
name=info(0)
if action<>"" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT �ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
rs.close
rs.open "select �ƶq,ID,�\��,�W�[�ƭ� from �׽m���~ where ���~�W='" & action & "' and �֦���=" & info(9) & " and �ƶq>0",conn
if rs.eof or rs.bof then
mess="�A�S���o�ت��~�I"
else
if int(rs("�ƶq"))<sl then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�A�����~�u��"&rs("�ƶq")&"�ӡA�ӧA�Q�A��"&sl&"��,�A�����~�����I');</script>"
	%>
	<script language="vbscript">
	location.href = "javascript:history.go(-1)"
	</script>
	<%
	response.end
end if
id=rs("ID")
gx=rs("�\��")
jjsj=int(rs("�W�[�ƭ�"))*sl
conn.execute "update �׽m���~ set �ƶq=�ƶq-"& sl &" where id=" & id
conn.execute "delete * from �׽m���~ where �ƶq<=0"
if gx="���H��" then
	conn.execute "update �Τ� set "& gx &"="& jjsj &",�ާ@�ɶ�=now() where id="&info(9)
else
	conn.execute "update �Τ� set "& gx &"="& gx &"+"& jjsj &" ,�ާ@�ɶ�=now() where id="&info(9)
end if
mess="�A��ΤF�Ī�"& action & sl &"�ӡA�W�[" & gx & jjsj & "�I�I"
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����</title></head>

<body background="back.gif" oncontextmenu=self.event.returnValue=false>
<table border="0" align="center" width="300" cellspacing="0" cellpadding="0">
<tr align="center">
<td width="300" height="30"><font
color="FF6600"><b>���򴣥�</b></font></td>
</tr>
<tr align="center">
<td>
<table width="260">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
</td>
</tr>
</table>
</td>
</tr>
<tr align="center">
<td width="500" height="30"><a href="javascript:location.replace('eat.asp');" onmouseover="window.status='��^';return true;" onmouseout="window.status='';return true;">��^�˳Ƥ@��</a></td>
</tr>
</table>
</body>