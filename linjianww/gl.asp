<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim tmprs
tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where times<3 and CDate(lasttime)<now()-10")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
user1=musers
tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where allvalue<=5 and CDate(lasttime)<now()-10")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
user2=musers
tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where CDate(lasttime)<now()-30")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers

%><title>�F�C����ƾڮw���@�{��</title><style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
<body background="0064.jpg">
<div align="center">
<p>�D�Τ�ƾڮw�G</p>
<table width="500" border="1" bordercolor="#000000" cellspacing="0" cellpadding="1">
<tr>
<td width="85%">10�Ѥ��e,�ȵn��2�����D�|�����G<font color="#0000FF"><b><%=user1%>��</b></font></td>
<td width="15%">
<div align="center"><a href="gl1.asp">�R��</a></div>
</td>
</tr>
<tr>
<td width="85%">10�Ѥ��e�w�I��allvalue�p��5���D�|�����G<b><font color="#0000FF"><%=user2%>��</font></b></td>
<td width="15%">
<div align="center"><a href="gl2.asp">�R��</a></div>
</td>
</tr>
<tr>
<td width="85%">
        <p>30�ѱq���n�����D�|�����G<font color="#0000FF"><b><%=user3%>��</b></font></p>
      </td>
<td width="15%">
<div align="center"><a href="gl3.asp">�R��</a></div>
</td>
</tr>
</table>
  <br>
  <p align="center"><BR>
    �b�o�̧A�i�H�d�ݨ��e�ƾڮw���ϥα��p�A<br>
    ��ܧR���N�i�H��o�@�ǥΤ�R�����A�R�������ܼƾڮw���Y�ާ@�I<br>
    �b�ڪ�����W�ڴN�γo�@�Ǥ�k���\��@��11MB���ƾڮw���Y��4MB�j�p�I<br>
    �b���ާ@�e��ĳ�ƥ���l���A�p�G�]���{�Ǿާ@�y������G�ڭ̤��t����d���I<br>
    ���:�O���ƾڮw�BBBS�ƾڮw�p�G�S�O�j���i�H��ܥέ�l����л\�w�Y�i�A<br>
    ��ĳ��o�Ǥ��ƥ�<BR>
</div> 