<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=Trim(Request.QueryString("id"))
Session("ljjh_inthechat")<>"1"
Select Case id
Case "000"
nl="  �藍�_�A�{�ǩҦb�ؿ����O�����ؿ��A�γ]�m�����~�AGlobal.asa �S���Q����C���{�ǻݭn�����ؿ�������I"
Case "001"
if session("nowinroom")<>"" then
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
ubb=UBound(onlinelist)
for i=1 to ubb step 6
if CStr(onlinelist(i+1))<>CStr(Session("ljjh_name")) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock

nl="�����f�Y�N�Q�����C<br><br>�y�����{������]�D�n���G<br>   �ѩ�����ǿ���D�A�ɭP�A�������P�A�Ⱦ��b�T�������L�k�ǻ��ƾڡA�t�αN�A�@���W�ɦ��_�}�F�s���F<br>   �A�I���F�����W�n�������^�n�������A�S�S�������������f�F<br>   �A�Q��X�F��ѫǡC<br>   �p�G�A�ϥΥ����f�i�J���|�X�{�o�Ӱ��D�A�ӥηs���f�i�J�N�X�{�o�Ӱ��D���ܡA���D�N�X�b�A�������G��u�X���s���f�L�k�~�ӤW�ŵ��f���ȡC<br><br>�ѨM��k�G<br>  ���������f�A���s��n��������J�Τ�W�M�K�X�i��n���C�p�G�A�O�Q��X��ѫǡA�h�A��~�ҥΪ��Τ�W�b 5 ����������ϥΡC�p�G�O�Ģ��ر��p�A�i�H�յ۲M���עӪ��{�ɤ��A�M�᭫�s�Ұʾ����C"
else
nl="�����f�Y�N�Q�����C<br><br>�y�����{������]�D�n���G<br>   �ѩ�����ǿ���D�A�ɭP�A�������P�A�Ⱦ��b�T�������L�k�ǻ��ƾڡA�t�αN�A�@���W�ɦ��_�}�F�s���F<br>   �A�I���F�����W�n�������^�n�������A�S�S�������������f�F<br>   �A�Q��X�F��ѫǡC<br>   �p�G�A�ϥΥ����f�i�J���|�X�{�o�Ӱ��D�A�ӥηs���f�i�J�N�X�{�o�Ӱ��D���ܡA���D�N�X�b�A�������G��u�X���s���f�L�k�~�ӤW�ŵ��f���ȡC<br><br>�ѨM��k�G<br>  ���������f�A���s��n��������J�Τ�W�M�K�X�i��n���C�p�G�A�O�Q��X��ѫǡA�h�A��~�ҥΪ��Τ�W�b 5 ����������ϥΡC�p�G�O�Ģ��ر��p�A�i�H�յ۲M���עӪ��{�ɤ��A�M�᭫�s�Ұʾ����C"
end if
Case "002"
nl="�����f�Y�N�Q�����C<br><br>��]�G<br>  �A�W�L120�����S���o���A����A�Ⱦ��t��A�t�Φ۰ʲM���A�ҥe�Ϊ��귽�C<br><br>�ѨM��k�G<br>  ���������f�A���s��n��������J�Τ�W�M�K�X�i��n���A�p�G�h���X�{�o�ر��p�i���F����N�I�������M��IE�̤U��IE�M�����v�O���n��B��N�i�H�i����F�I"
Case else
nl="  �藍�_�A�ӥX���������Q�n�O�C"
End Select
Session.Abandon
nl="<p>" & nl & "</p>"%><html>
<head>
<title>�X������</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=<%=chatroombgcolor%> background=<%=chatroombgimage%>>
<script LANGUAGE="JavaScript">if(window!=window.top)top.location.href=location.href;</script>
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="E0E0E0">
<tr>
<td>
<table border="0" bgcolor="#FF0099" cellspacing="0" cellpadding="2" width="350">
<tr>
<td width="342"><font color="FFFFFF"> �X������</font></td>
<td width="18">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="����"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr>
<td width="59" align="center" valign="top"><font face="Wingdings" color="#FF0000" style="font-size:32pt">L</font></td>
<td width="269">
<%=nl%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value=" �T �w " onclick=javascript:window.close()>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html> 