<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="const.inc"-->
<%
if datediff("s",session("refresh"),now())<5 then
response.write "穝Ω计びе,叫篊翴!"
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
rs.open "Select 猌,ず,ю阑,ň眘,砰,緔,笵紈,猭翴 from ノめ where id="&info(9),conn
wgj=rs("猌")
nlj=rs("ず")
gjj=rs("ю阑")
fyj=rs("ň眘")
tlj=rs("砰")
mlj=rs("緔")
ddj=rs("笵紈")
flj=rs("猭翴")
rs.close
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>猭竟</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 穝灿砰; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 穝灿砰; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 穝灿砰; font-size: 9pt; font-style: bold;}
</style>
<HTML>
<HEAD>
<title>篈</title>
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
rs=conn.execute("Select ﹎,闹,戮穨,畍撑,皌案,,跋,穦单,穦ら戳,腳,炳计,ō,盉,,刽,单,grade,蝗ㄢ,蹿,猭,猭翴,借,獁ě翴计,玂臔,忌ě丁,穦禣,砰,ず,猌,ю阑,ň眘,緔,笵紈,a1,a2,a3,a4,a5,a6,b1,b2,b3,b4,b5,b6,幅 from ノめ where ﹎='"&info(0)&"'")
if not rs.eof then
%>
闹<%=rs("闹")%><br>
戮穨<%=rs("戮穨")%><br>
畍撑<%=rs("畍撑")%><br>
<%=rs("")%><br>
幅<%=rs("幅")%><br>
跋<%=rs("跋")%><br>
<%if rs("穦单")>0 then%>
穦<%=rs("穦单")%><br>
穦ら戳<%=rs("穦ら戳")%><br>
<%end if%>
腳<%=rs("腳")%><br>
炳<%=rs("炳计")%><br>
ō<%=rs("ō")%><br>
<%if info(1)="╧" then %>
ρ盋
<%else%>
ρそ
<%end if%>
<%=rs("皌案")%><br>
<%if info(1)="╧" then %>
ρ盋
<%else%>
ρそ
<%end if%>
<%=rs("盉")%><br>
<%=rs("")%><br>
单<%=rs("单")%><br>
恨瞶<%=rs("grade")%><br>
蝗ㄢ<%=rs("蝗ㄢ")%><br>
蹿<%=rs("蹿")%><br>
猭<%=rs("猭")%><br>
腳瞺<%=rs("猭翴")%><br>
戈借<%=rs("借")%><br>
ě翴
<%if rs("grade")=3 and rs("ō")="臔猭" then%>
<%=rs("獁ě翴计")-int(rs("獁ě翴计")/500)*500%>/500<br>
<%else%>
<%=rs("獁ě翴计")-int(rs("獁ě翴计")/1000)*1000%>/1000<br>
<%end if%>
ě
<%if rs("grade")=3 and rs("ō")="臔猭" then%>
<%=int(rs("獁ě翴计")/500)%><br>
<%else%>
<%=int(rs("獁ě翴计")/1000)%><br>
<%end if%>
絤
<%if rs("玂臔")=true then%>
絤玂臔<br>
<%else%>
⊿Τ玂臔<br>
<%end if%>
<%
sj=DateDiff("n",rs("忌ě丁"),now())
if sj<=60 then%>
临逞<%=60-sj%>だ<br>
<%end if%>
腹
<%
if rs("单")<5  then response.write("ㄓ笵")
if rs("单")>=5 and rs("单")<10  then response.write("打︽")
if rs("单")>=10 and rs("单")<15  then response.write("ΤΘ碞")
if rs("单")>=15 and rs("单")<20  then response.write("羘陪划")
if rs("单")>=20 and rs("单")<35  then response.write("︽卖打")
if rs("单")>=35 and rs("单")<45  then response.write("獿")
if rs("单")>=45 and rs("单")<55  then response.write("打糃")
if rs("单")>=55 and rs("单")<65  then response.write("籇ぱ")
if rs("单")>=65 and rs("单")<75  then response.write("冻笴秤")
if rs("单")>=75 and rs("单")<80  then response.write("笵")
if rs("单")>=80 and rs("单")<85  then response.write("")
if rs("单")>=85 and rs("单")<90  then response.write("")
if rs("单")>=90 and rs("单")<95  then response.write("伐")
if rs("单")>=95 and rs("单")<100  then response.write("")
if rs("单")>=100 then response.write("糃")%>
<br>
刽<%=rs("刽")%><br>
穦禣<%=rs("穦禣")%>じ<br>
砰
<img height=12
<%if int((rs("砰")/(rs("单")*1500+5260+tlj))*100)<100 then
abc=int(int((rs("砰")/(rs("单")*1500+5260+tlj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("砰")/(rs("单")*1500+5260+tlj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("砰")/(rs("单")*1500+5260+tlj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("砰")%>/<%=rs("单")*1500+5260+tlj%><br>
ず
<img height=12
<%if int((rs("ず")/(rs("单")*640+2000+nlj))*100)<100 then
abc=int(int((rs("ず")/(rs("单")*640+2000+nlj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("ず")/(rs("单")*640+2000+nlj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("ず")/(rs("单")*640+2000+nlj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("ず")%>/<%=rs("单")*640+2000+nlj%><br>
猌
<img height=12
<%if int((rs("猌")/(rs("单")*1280+3800+wgj))*100)<100 then
abc=int(int((rs("猌")/(rs("单")*1280+3800+wgj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("猌")/(rs("单")*1280+3800+wgj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("猌")/(rs("单")*1280+3800+wgj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("猌")%>/<%=rs("单")*1280+3800+wgj%><br>
ю阑/肂ю阑
<img height=12
<%if int(rs("单")*150)<100 then
abc=int(rs("单")*150)/25+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int(rs("单")*150)%>><img height=12 src="img/5.gif" width=8>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=int(rs("单")*150)%>/<%=rs("a1")+rs("a2")+rs("a3")+rs("a4")+rs("a5")+rs("a6")%><br>
ň眘/肂ň眘
<img height=12
<%if int(rs("单")*200)<100 then
abc=int(rs("单")*200)/25+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int(rs("单")*200)%>><img height=12 src="img/5.gif" width=<%=100-int(rs("单")*200)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=int(rs("单")*200)%>/<%=rs("b1")+rs("b2")+rs("b3")+rs("b4")+rs("b5")+rs("b6")%><br>
緔
<img height=12
<%if int((rs("緔")/(rs("单")*1120+2100+mlj))*100)<100 then
abc=int(int((rs("緔")/(rs("单")*1120+2100+mlj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("緔")/(rs("单")*1120+2100+mlj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("緔")/(rs("单")*1120+2100+mlj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("緔")%>/<%=rs("单")*1120+2100+mlj%><br>
笵紈
<img height=12
<%if int((rs("笵紈")/(rs("单")*1440+2200+ddj))*100)<100 then
abc=int(int((rs("笵紈")/(rs("单")*1440+2200+ddj))*100)/25)+1
fi="img/"&abc&".gif"
%>
src="<%=fi%>" width=<%=int((rs("笵紈")/(rs("单")*1440+2200+ddj))*100)%>><img height=12 src="img/5.gif" width=<%=100-int((rs("笵紈")/(rs("单")*1440+2200+ddj))*100)%>>
<%else%>
src="img/4.gif" width=100>
<%end if%>
<br><%=rs("笵紈")%>/<%=rs("单")*1440+2200+ddj%><br>
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