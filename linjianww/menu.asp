<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
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
td, div, form, option, p, td, br{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: �s�ө���; font-size: 9pt; font-style: bold;}
</style>
<title>�����</title>
<base target="main">
<link rel="stylesheet" >
</head>
<body link="#FF6600" vlink="#FF6600" text="#FF6600" alink="#FF6600" bgcolor=FFFFFF leftmargin="0" topmargin="0">
<table  border="1" width="100%"  cellspacing=0 cellpadding=3  align="center" class="cp">
<tr><td><a href="managelist.asp" target="txt">�޲z�O��</a></td></tr>
<tr><td><a href="m.asp" target="txt">�K�[����</a></td></tr>
<tr><td><a href="adminmpp.asp" target="txt">�����޲z</a></td></tr>
<tr><td><a href="fine.asp" target="txt">�Τ�޲z</a></td></tr>
<tr><td><a href="adminaa.asp" target="txt">���[�޲z</a></td></tr>
<tr><td><a href="seeuser.asp" target="txt">����d��</a></td></tr>
<tr><td><a href="ljyamen.asp" target="txt">�x���޲z</a></td></tr>
<tr><td><a href="ljroom.asp" target="txt">�ж��޲z</a></td></tr>
<tr><td><a href="hyadmin.asp" target="txt">�|���޲z</a></td></tr>
<tr><td><a href="fawupin.asp" target="txt">���~�o�e</a></td></tr>
<tr><td><a href="ljmoney.asp" target="txt">�u��޲z</a></td></tr>
<tr><td><a href="ljsetwg.asp" target="txt">�Z�\�޲z</a></td></tr>
<tr><td><a href="binqi.asp" target="txt">�L���޲z</a></td></tr>
<tr><td><a href="yaopu.asp" target="txt">�ľQ�޲z</a></td></tr>
<tr><td><a href="ljsheep.asp" target="txt">�d���޲z</a></td></tr>
<tr><td><a href="wpmoney.asp" target="txt">�����վ�</a></td></tr>
<tr><td><a href="carda.asp" target="txt">�d���޲z</a></td></tr>
<tr><td><a href="fangwu.asp" target="txt">�Ыκ޲z</a></td></tr>
<tr><td><a href="../sy/manage.asp" target="txt">�ӭ޺޲z</a></td></tr>
<tr><td><a href="modifygg.asp" target="txt">�s�i�ק�</a></td></tr>
<tr><td><a href="modifywb.asp" target="txt">���a�p��</a></td></tr>
<tr><td><a href="ljbbs.asp" target="txt">�׾º޲z</a></td></tr>
<tr><td><a href="ip.asp" target="txt">IP�޲z</a></td></tr>
<tr><td><a href="sendw.asp" target="txt">�ذe���~</a></td></tr>
<tr><td><a href="manaccquerygrade.asp" target="txt">�d�ݽ㸹</a></td></tr>
<tr><td><a href="manacc.asp" target="txt">�㸹�޲z</a></td></tr>
<tr><td><a href="sqlcomm.asp" target="txt">��ߢں޲z</a></td></tr>
<tr><td><a href="userok.asp" target="txt">�䧽�޲z</a></td></tr>
<tr><td><a href="../linjiankkk/compmdb.asp" target="txt">���Y�ƾڮw</a></td></tr>
<tr><td><a href="gl.asp" target="txt">�ƾں޲z</a></td></tr>
<tr><td><a href="ljother.asp" target="txt">�����޲z</a></td></tr>
<tr><td><a href="ljgwu.asp" target="txt">�Ǫ��޲z</a></td></tr>
<tr><td><a href="http://add.yahoo.com/fast/add?68577218+HK" target="txt">�n������</a></td></tr>
<tr><td><a target="newbook" href="admin_news/newslist.asp?type=3">�o�G���i</a><p>
  <a href="http://t0306.somee.com/new_news/pwd.asp">�s���i�o</a></td></tr>
</table>
</body>
</html>