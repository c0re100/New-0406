<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>�b�u���</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: �s�ө���; font-size: 9pt; font-style: bold;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" text="#FFFFFF" link="#FFFFFF" topmargin="0" vlink="#FFFFFF" leftmargin="0"  background=bk.jpg>
<div align="center">

<table style=filter:shadow(color=#CC99FF)    border="1" cellspacing="0" cellpadding="1" width=95% bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center"  bgcolor=#006BA5>
<tr><td bgcolor="#EEFFEE" height="62" ><div align="center"><font color="#FF6633"> �b�u��� </font></div></td></tr>
<tr><td>
<form  method=POST  action='duju/duju001.asp'><div align="center">
<br><input title="�n�D�G�԰��Ťj��30�šB�s��2000�U�I" type=submit  name="�Ȩ⧽" value=�Ȩ⧽�@��>
<br><input title="�n�D�G�԰��Ťj��30�šB�k�O5000�I�I" type=submit  name="�k�O��" value=�k�O���@��>
<br><input title="�n�D�G�԰��Ťj��30�šB�w�I��3000�I�I" type=submit  name="�w�I��" value=�w�I���@��>
<br><input title="�n�D�G�԰��Ťj��30�šB������4000�ӡI" type=submit  name="������" value=�������@��>
<br><input title="�[�ܨ�L���a���֤U�`�I" type=submit  name="�[��" value=�U�䧽�[��>
</div>
</form>
</td></tr>
<tr><td><form  method=POST  action='duju/duju002.asp'><div align="center"><input title="�Ȩ�䧽�I" type=submit  name="�Ȩ⧽" value=�Ȩ⧽�U�`></div></form></td></tr>
<tr><td><form  method=POST  action='duju/duju003.asp'><div align="center"><input title="�k�O�䧽�I" type=submit  name="�k�O��" value=�k�O���U�`></div></form></td></tr>
<tr><td><form  method=POST  action='duju/duju004.asp'><div align="center"><input title="�w�I�䧽�I" type=submit  name="�w�I��" value=�w�I���U�`></div></form></td></tr>
<tr><td><form  method=POST  action='duju/duju005.asp'><div align="center"><input title="�����I" type=submit  name="������" value=�������U�`></div></form></td></tr>
<tr><td><form  method=POST  action='duju/DPK-XP.ASP'><div align="center"><input title="���J�P��" type=submit  name="���J�P��" value=���J�P��></div></form></td></tr>
<tr><td><form  method=POST  action='duju/DMJ-XP.ASP'><div align="center"><input title="�±N�P��" type=submit  name="�±N�P��" value=�±N�P��></div></form></td></tr>
</table>

</div>
</body>
</html>
