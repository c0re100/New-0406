<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
time1=request.form("time")+0
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if time1=0 then
mess="�п�ܥ��T���ɶ��I"
elseif  time1<0 then
mess="�п�ܥ��T���ɶ��I"
elseif  time1>24 then
mess="�п�ܥ��T���ɶ��I"
else
Jname=info(0)
Jpass=request.cookies("pass")
yin=abs((time1)*10)
shi=0.0416*time1
rs.open "select �Ȩ� from �Τ� where id=" & info(9) & "  and ���A<>'��' and ���A<>'��'",conn
if rs.eof or rs.bof then
mess=Jname & "�A�A����ާ@�I"
else
if rs("�Ȩ�")<yin then
mess=Jname & "�A�A���Ȩ⤣���I"
else
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & ",��O=��O+"& yin &",���O=���O+"&yin&",�n��=now()+" & shi & ",���A='�v' where id="&info(9)
mess=Jname & "�A�A�w�g���\����ȴ̥𮧤F�A�Цb" &time1& "�p�ɫ�ϥθӱb���I"
session.Abandon
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.htm"
</script>
<%
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����ȴ�</title></head>

<body bgcolor="#000000" background="../../images/8.jpg" text="#000000" link="#0000FF">
<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="pub.asp">��^</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>