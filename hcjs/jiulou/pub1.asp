<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
jiu=request.form("jiu")
Jname=info(0)
select case jiu
case "lao"
tili=50
yin=250
case "wu"
tili=60
yin=300
case "du"
tili=70
yin=350
case "mao"
tili=80
yin=400
end select
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ�,�Z�\ from �Τ� where id=" & info(9) & " and ���A<>'��' and ���A<>'��' ",conn
if rs.eof or rs.bof then
mess=Jname & "�A�A����ާ@�I"
else
if rs("�Ȩ�")<yin then
mess=Jname & "�A�A���Ȩ⤣���I"
else
wugong=rs("�Z�\")
jiuliang=int((wugong)/100)
if jiuliang>1 then
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & ",��O=��O+"& tili &" where id="&info(9)
mess=Jname & "�n�s�q!�A�ܧ��@�M,�������ٹD:�n�s,�n�s!"
else	
shi=0.0416*1
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-" & yin & ",��O=��O+"& tili &",�n��=now()+" & shi & ",���A='�v' where id="&info(9)
mess=Jname & "�A�ѩ�z�s�q����,�w�b��W�N�εۤF,�X�ӥ�p�N�z���ȴ̥h�𮧤F�A�Цb1�p�ɫ�ϥθӱb���I"
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script>
confirm('<%=mess%>')
top.menu.location.href="../../index.asp"
</script>
<%
end if
end if
end if
%>

<head>
<style>td           { font-size: 14 }
</style>
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
<p align="center"><a href="pub.asp">��^</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body>

