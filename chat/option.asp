<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")%>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
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
<script language=javascript>
function exitchat(){
if(confirm('�A�T�H�h�X�F�C����ܡH')){
parent.resfrm.location.href='exitlt.asp';
}}
function MM_openBrWindow(theURL,winName,features) { //v2.0
window.open(theURL,winName,features);
}
</script>
<title>�\���</title></head>
<body oncontextmenu=self.event.returnValue=false leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background=bk.jpg>
<div align=center>
<table border=0 cellspacing=2 cellpadding="2" width="100%">
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���A" onClick="window.open('../yamen/seeme.asp','zhuangti','scrollbars=no,resizable=no,width=450,height=320')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���~"  onClick="window.open('../hcjs/jhzb/head.asp','wupin','scrollbars=yes,resizable=yes,width=750,height=452')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�ľQ" onClick="window.open('../hcjs/jhjs/yaopu.asp')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="��Q" onClick="window.open('../hcjs/jhjs/dan.asp')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�L��" onClick="window.open('../hcjs/jhjs/binqi.asp')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�c��" onClick="window.open('../yamen/listlao.asp')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�R��" onClick="window.open('../yamen/minan.asp')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���" onClick="window.open('../yamen/regmodi.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="��O" onClick="window.open('../Diary/Diary.asp')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�B��" onClick="window.open('../hunyin/yuelao.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�j��" onClick="window.open('../jl/Top.htm')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�ѥ�" onClick="window.open('../gupiao/index.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="����" onClick="window.open('../hcjs/yilao.asp')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���" onClick="window.open('../xianpian/main.asp?action=search','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���]" onClick="window.open('../qmg/qmg.htm')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�Z�\" onClick="window.open('../jhmp/setwg.asp','zhuangti','scrollbars=yes,resizable=no,width=450,height=290')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="��ë" onClick="window.open('../jhqy/wish.asp','jhqy','scrollbars=yes,resizable=yes,width=750,height=352')" ></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���_" onClick="window.open('../baozang/diaoyu.asp')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="����" onClick="window.open('../bianfu/daying.asp','bianfu','scrollbars=no,resizable=no,width=750,height=352')" title="�������A�o�_���I�I"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="����" onClick="window.open('../diaoyu/diaoyu.asp')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�ɰ�" onClick="window.open('../bet/sama/horse.asp','ljjhhorse','scrollbars=yes,resizable=yes')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="��W" onClick="window.open('../gmhx/gmhx.asp','gmhx','scrollbars=yes,resizable=yes')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="�{��" onClick="window.open('../jhmp/MYTUDI.ASP','MYTUDI','scrollbars=yes,resizable=yes,width=750,height=352')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���I" onClick="window.open('../jhmp/MYFRIEND.ASP','MYFRIEND','scrollbars=yes,resizable=yes,width=750,height=352')"></a></div></td>
</tr>
<tr>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="���" onClick="window.open('../work/famu/famuMAIN.ASP','famu','scrollbars=yes,resizable=yes,width=450,height=352')"></a></td>
<td><div align="center"><INPUT TYPE="BUTTON" VALUE="RPG" onClick="window.open('../rpg/index.ASP','rpg','scrollbars=yes,resizable=yes,width=450,height=352')"></a></div></td>
</tr>
</table>
</div>
</body>

