<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="const.inc"-->
<%
if datediff("s",session("refresh"),now())<5 then
response.write "�A��s���Ƥӧ�,�Щ�C�I!"
response.end
else
session("refresh")=now()
end if
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
on error resume next
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select �Z�\�[,���O�[,�����[,���s�[,��O�[,�y�O�[,�D�w�[,�k�O�I from �Τ� where id="&info(9),conn
wgj=rs("�Z�\�[")
nlj=rs("���O�[")
gjj=rs("�����[")
fyj=rs("���s�[")
tlj=rs("��O�[")
mlj=rs("�y�O�[")
ddj=rs("�D�w�[")
flj=rs("�k�O�I")
rs.close
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>�k��</title>
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
<HTML>
<HEAD>
<title>���A</title>
<style type=text/css>
<!--   td {  line-height: 140%}
--></style>
<script language="JavaScript">
<!--
function selectwho(list){
parent.f2.document.forms[0].towho.options[0].value=list;
parent.f2.document.forms[0].towho.options[0].text=list;
parent.f2.document.forms[0].saystemp.focus();
}
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></HEAD>
<BODY bgcolor="#ffffff" background="bk.jpg">
<TABLE border="0" cellpadding="0" cellspacing="0" width="142">
<TR> 
<TD align="center" height="10" width="20" valign="top"><IMG src="./fq/tb-a1.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a6.gif" height="10" width="102"></TD>
<TD height="10" align="center" valign="top" width="57"><IMG src="./fq/tb-a7.gif" width="20" height="20" border="0"></TD>
</TR>
<TR> 
<TD width="20" background="./fq/tb-a2.gif" height="80"></TD>
<TD width="102" height="80" align="center" valign="top" background="./fq/tb-a3.gif">
<table border="0" cellspacing="0" cellpadding="1" width="102" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="287" >
<div align="center"><tr><td>
<%
rs=conn.execute("Select �m�W,�~��,¾�~,�v��,�t��,����,�a��,�|������,�|�����,�_��,���H��,����,�G�B,�p��,����,����,grade,�Ȩ�,�s��,�k�O,�k�O�I,����,�w���I��,�O�@,�ɨ��ɶ�,�|���O,��O,���O,�Z�\,����,���s,�y�O,�D�w,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6,�P�� from �Τ� where �m�W='"&info(0)&"'")
if not rs.eof then
%>
�~�֡G<%=rs("�~��")%><br>
¾�~�G<%=rs("¾�~")%><br>
�v�šG<%=rs("�v��")%><br>
�����G<%=rs("����")%><br>
�P���G<%=rs("�P��")%><br>
�a�ϡG<%=rs("�a��")%><br>
<%if rs("�|������")>0 then%>
�|���G<%=rs("�|������")%><br>
�|������G<%=rs("�|�����")%><br>
<%end if%>
�_���G<%=rs("�_��")%><br>
���H�G<%=rs("���H��")%><br>
�����G<%=rs("����")%><br>
<%if info(1)="�k" then %>
�j�ѱC�G
<%else%>
�j�Ѥ��G
<%end if%>
<%=rs("�t��")%><br>
<%if info(1)="�k" then %>
�p�ѱC�G
<%else%>
�p�Ѥ��G
<%end if%>
<%=rs("�G�B")%><br>
�p�ġG<%=rs("�p��")%><br>
���šG<%=rs("����")%><br>
�޲z�G<%=rs("grade")%><br>
�Ȩ�G<%=rs("�Ȩ�")%><br>
�s�ڡG<%=rs("�s��")%><br>
�k�O�G<%=rs("�k�O")%><br>
�_�~�G<%=rs("�k�O�I")%><br>
���G<%=rs("����")%><br>
���I�G
<%if rs("grade")=3 and rs("����")="�@�k" then%>
<%=rs("�w���I��")-int(rs("�w���I��")/500)*500%>/500<br>
<%else%>
<%=rs("�w���I��")-int(rs("�w���I��")/1000)*1000%>/1000<br>
<%end if%>
���l�G
<%if rs("grade")=3 and rs("����")="�@�k" then%>
<%=int(rs("�w���I��")/500)%><br>
<%else%>
<%=int(rs("�w���I��")/1000)%><br>
<%end if%>
�m�\�G
<%if rs("�O�@")=true then%>
�m�\�O�@<br>
<%else%>
�S���O�@<br>
<%end if%>
<%
sj=DateDiff("n",rs("�ɨ��ɶ�"),now())
if sj<=60 then%>
�¤O�ٳѡG<%=60-sj%>��<br>
<%end if%>
�W���G
<%
if rs("����")<5  then response.write("��ӥE�D")
if rs("����")>=5 and rs("����")<10  then response.write("������")
if rs("����")>=10 and rs("����")<15  then response.write("�p�����N")
if rs("����")>=15 and rs("����")<20  then response.write("�n�W�㻮")
if rs("����")>=20 and rs("����")<35  then response.write("��������")
if rs("����")>=35 and rs("����")<45  then response.write("�@�N�j�L")
if rs("����")>=45 and rs("����")<55  then response.write("����C��")
if rs("����")>=55 and rs("����")<65  then response.write("�D�W�ѤU")
if rs("����")>=65 and rs("����")<75  then response.write("���C�P��")
if rs("����")>=75 and rs("����")<80  then response.write("�w�J�P�D")
if rs("����")>=80 and rs("����")<85  then response.write("�p�P")
if rs("����")>=85 and rs("����")<90  then response.write("�j�P")
if rs("����")>=90 and rs("����")<95  then response.write("���Ҥj�P")
if rs("����")>=95 and rs("����")<100  then response.write("�P�H")
if rs("����")>=100 then response.write("�C�P")%>
<br>
�����G<%=rs("����")%>��<br>
�|�O�G<%=rs("�|���O")%>��<br>
��O�G
<img height=12
<%if int((rs("��O")/(rs("����")*1500+5260+tlj))*100)<100 then
abc=int(int((rs("��O")/(rs("����")*1500+5260+tlj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("��O")/(rs("����")*1500+5260+tlj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("��O")/(rs("����")*1500+5260+tlj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("��O")%>/<%=rs("����")*1500+5260+tlj%><br>
���O�G
<img height=12
<%if int((rs("���O")/(rs("����")*640+2000+nlj))*100)<100 then
abc=int(int((rs("���O")/(rs("����")*640+2000+nlj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("���O")/(rs("����")*640+2000+nlj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("���O")/(rs("����")*640+2000+nlj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("���O")%>/<%=rs("����")*640+2000+nlj%><br>
�Z�\�G
<img height=12
<%if int((rs("�Z�\")/(rs("����")*1280+3800+wgj))*100)<100 then
abc=int(int((rs("�Z�\")/(rs("����")*1280+3800+wgj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("�Z�\")/(rs("����")*1280+3800+wgj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("�Z�\")/(rs("����")*1280+3800+wgj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("�Z�\")%>/<%=rs("����")*1280+3800+wgj%><br>
����/�B�~�����O�G
<img height=12
<%if int(rs("����")*150)<100 then
abc=int(rs("����")*150)/25+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int(rs("����")*150)%>><img height=12 src="img/5.gif" width=8>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=int(rs("����")*150)%>/<%=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")%><br>
���s/�B�~���s�O�G
<img height=12
<%if int(rs("����")*200)<100 then
abc=int(rs("����")*200)/25+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int(rs("����")*200)%>><img height=12 src="img/5.gif" width=<%=100-int(rs("����")*200)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=int(rs("����")*200)%>/<%=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")%><br>
�y�O�G
<img height=12
<%if int((rs("�y�O")/(rs("����")*1120+2100+mlj))*100)<100 then
abc=int(int((rs("�y�O")/(rs("����")*1120+2100+mlj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("�y�O")/(rs("����")*1120+2100+mlj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("�y�O")/(rs("����")*1120+2100+mlj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("�y�O")%>/<%=rs("����")*1120+2100+mlj%><br>
�D�w�G
<img height=12
<%if int((rs("�D�w")/(rs("����")*1440+2200+ddj))*100)<100 then
abc=int(int((rs("�D�w")/(rs("����")*1440+2200+ddj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("�D�w")/(rs("����")*1440+2200+ddj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("�D�w")/(rs("����")*1440+2200+ddj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("�D�w")%>/<%=rs("����")*1440+2200+ddj%><br>
</td>
</tr>
</table></TD>
<TD background="./fq/tb-a4.gif" width="57" height="80"></TD>
</TR>
<TR> 
<TD width="20" height="10"><IMG src="./fq/tb-a8.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a5.gif" width="102" height="10"></TD>
<TD width="57" height="10"><IMG src="./fq/tb-a10.gif" width="20" height="20" border="0"></TD>
</TR></form></TABLE></BODY></HTML>
<%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>