<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
nowdate=cstr(date())
bkmn=Request.Form("money")
bkop=Request.Form("op")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if isnumeric(bkmn) then
	if bkmn>1 then
		if bkop="����" then	
conn.Execute "update �Τ� set �Ȩ�=�Ȩ�+"&bkmn&",�s��=�s��-"&bkmn&" where �m�W='"&username&"' and �s��>="&bkmn
msg="�z�q�����̦��\���X<font color=red>["&bkmn&"]</font>��Ȥl�A�u�A�ٵ��I��C�I"
end if
		if bkop="�s��" then 
conn.Execute "update �Τ� set �Ȩ�=�Ȩ�-"&bkmn&",�s��=�s��+"&bkmn&" where �m�W='"&username&"' and �Ȩ�>="&bkmn
msg="�z�b�����̦��\�s�J<font color=red>["&bkmn&"]</font>��Ȥl�A�h�h�s���ڡA�N�ӭn�i�a�k�f���ڡI"
end if
set rs=nothing
conn.close
set conn=nothing		
		
	else
		msg="�}�������a�A�A�쩳�O�Q�s�٬O�Q���I"
	end if	
else	
	msg="�зǽT��J���B�I"
end if	

%><head><link rel="stylesheet" href="../style.css"></head>

<title>�F�C�������</title>
<body bgcolor='<%=bgcolor%>' background='../bg.gif' oncontextmenu=self.event.returnValue=false>
<p align=center>�F�C������� <div align="center"> <table width="75%" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000"> 
<tr> <td bgcolor="#FFFFFF"> <p> <%=msg%> <p align=center><a href="bankok.asp" onMouseOver="window.status='��^����';return true;" onMouseOut="window.status='';return true;"><b></b></a></p></td></tr> 
</table></div><p align="center"><a href="bankok.asp" onmouseover="window.status='��^����';return true;" onmouseout="window.status='';return true;"><b>��^</b></a> 
<p align="center">&nbsp;
</body>