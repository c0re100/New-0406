<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"hla" then
response.redirect "index.asp"
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �Τ� where id="&info(9)
Set Rs=conn.Execute(sql)
wgj=(rs("����")*1280+3800)+rs("�Z�\�[")
wga=rs("�Z�\")
if wga>=wgj then
sql="update �Τ� set �Ȩ�=�Ȩ�+2000000,����=����+10000,�Z�\='"& wgj &"' where id="&info(9)
conn.Execute(sql)
else
sql="update �Τ� set �Ȩ�=�Ȩ�+2000000,����=����+10000,�Z�\=�Z�\+200000 where id="&info(9)
conn.Execute(sql)
end if
if rs("�Z�\")>=wgj then 
sql="update �Τ� set �Z�\='"& wgj &"' where id="&info(9)
conn.Execute(sql)
end if
mess="<font color=FFCFCF>"&name&"</font>�j�L�i���ؤs�����A���\���˵ؤs���j����A���}�F�_�w,���o�F�ѤU�L�Ī��Z�L���y�I"
set rs=nothing
conn.close
set conn=nothing
Session("iq")=""
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=FFCFCF>�ؤs�ǨӪ�����:"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>

<html>

<head>
<title>�A�ש����F�W�t�D�Ѫ��_��</title>
<LINK href="../../chat/READONLY/Style.css" rel=stylesheet>
</head>
<body background="images/Bg.gif" oncontextmenu=self.event.returnValue=false >
<div align="center"> <table border="0" width="600"> <tr> <td width="100%"> <p align="left" style="line-height: 200%; margin: 20"><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
�ש��F�W�t�D�Ѫ��_�äF�A�i�O��A�Τ��O�_�}�_�ê����O�A�o�o�{�èS������ѤU�L�Ī��Z�L���y�A���O�s�}���۾��W�s����R���g�۳o�ˤ@��r�G</b></td></tr> 
<tr> <td width="100%"><img border="0" src="finish.gif"></td></tr> <tr> <td width="100%"> 
<p align="center" style="line-height: 200%; margin: 20"><b><font color="#000000" size="3">�ѤU�L�Ī����y��²��</font></b><font color="#000000" size="3">�A<b>���N�O�i��M�۫H</b></font></p><p align="left" style="line-height: 200%; margin: 20"><b><font color="#FF0000">&nbsp;</font><font color="#000000">&nbsp;&nbsp; 
��ӿW�t�D�Ѳץͥ��D�o�@�Ѫ���]�N�O�L���i��M�۫H�r�I��ӧڭ̪��i��M�۫H�N�O�ѤU�L�Ī����y�r�I�]�b�s�}�̧��</font><font color="#FF0000">50000000</font><font color="#000000">��Ȥl�A���۾��W�W�t�D�ѯu�񪺱ҵo�A���W�[</font><font color="#FF0000">10000</font><font color="#000000">�I�A�Z�\�W�[�F</font><font color="#FF0000">200000</font><font color="#000000">�I�^</font></b> 
</td></tr> </table></div>
</html>
