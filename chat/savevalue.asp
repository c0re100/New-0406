<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(useronlinename," "&info(0)&" ")=0 then Response.Redirect "chaterr.asp?id=001"
maxtimeout=120
lasttime=Session("ljjh_lasttime")
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
if DateDiff("n",lasttime,sj)>maxtimeout then
Response.Write "<script>top.location.href='nosaytimeout.asp';</script>"
Response.End
end if
wbname=""
wbpf=1
ip=Request.ServerVariables("REMOTE_ADDR")
If ip = "" Then ip = Request.ServerVariables("REMOTE_ADDR")
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../wb/ljjmwb.asa")
Conn.Open connstr
rs.open "SELECT barname FROM bar WHERE qdlm=true and ip='"&ip&"'",conn
if Not(rs.Eof and rs.Bof) then
wbname=rs("barname")
wbpf=2
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �D�|���w�I,�|���w�I1,�|���w�I2,�|���w�I3,�|���w�I4,�|���w�I5,�|���w�I6,�|���w�I7,�|���w�I8,�|���w�I9,�|���w�I10 FROM �奻 ",conn
pao0=rs("�D�|���w�I")
pao1=rs("�|���w�I1")
pao2=rs("�|���w�I2")
pao3=rs("�|���w�I3")
pao4=rs("�|���w�I4")
pao5=rs("�|���w�I5")
pao6=rs("�|���w�I6")
pao7=rs("�|���w�I7")
pao8=rs("�|���w�I8")
pao9=rs("�|���w�I9")
pao10=rs("�|���w�I10")
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))=CStr(info(0)) then
addvalue=DateDiff("n",onlinelist(i+5),sj)*wbpf
addvalue1=DateDiff("n",onlinelist(i+5),sj)*2*wbpf
presj=onlinelist(i+5)
onlinelist(i+5)=sj
end if
next
Application("ljjh_onlinelist"&session("nowinroom"))=onlinelist
Application.UnLock
Session("ljjh_savetime")=now()
ip=Request.ServerVariables("REMOTE_ADDR")
sfwg=1
if info(7)<>"" and info(7)<>"�L" then
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(7))&" ")=0 then
sfwg=1
else
sfwg=2
end if
end if
jhhy=pao0*10
Select Case info(4)
case 1
jhhy=pao1*10
case 2
jhhy=pao2*10
case 3
jhhy=pao3*10
case 4
jhhy=pao4*10
case 5
jhhy=pao5*10
case 6
jhhy=pao6*10
case 7
jhhy=pao7*10
case 8
jhhy=pao8*10
case 9
jhhy=pao9*10
case 10
jhhy=pao10*10
case else
jhhy=pao0*10
end select
if session("nowinroom")=0 then 
jhhy=jhhy*2*10
end if
rs.close
sql="SELECT ���O,�Ȩ�,�|���O,����,�Z�\,���H��,�ާ@�ɶ�,�k�O,�t��,grade,allvalue,�w���I��,mvalue,lasttime,lastip FROM �Τ� WHERE id="&info(9)
rs.open sql,conn,1,3
killsj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if killsj>30 and rs("���H��")>=10 then
conn.execute "update �Τ� set ���H��=���H��-1,�ާ@�ɶ�=now() where id="&info(9)
end if

if Not(rs.Eof and rs.Bof) then
rs("�Ȩ�")=rs("�Ȩ�")+int(addvalue1*jhhy*100)
rs("�|���O")=rs("�|���O")+int(addvalue1*jhhy*0.1)
rs("����")=rs("����")+int(addvalue1*jhhy*0.1)
rs("���O")=rs("���O")+int(addvalue1*sfwg*jhhy*10)
rs("�Z�\")=rs("�Z�\")+int(addvalue1/2*sfwg*jhhy*2)
rs("�k�O")=rs("�k�O")+int(addvalue1/2*jhhy*2)
bl=rs("�t��")
grade=rs("grade")
rs("allvalue")=rs("allvalue")+int(addvalue*jhhy*9)
rs("�w���I��")=rs("�w���I��")+addvalue
if DateDiff("m",presj,sj)=0 then
rs("mvalue")=rs("mvalue")+int(addvalue*jhhy*3)
else
rs("mvalue")=addvalue
end if
rs("lasttime")=sj
rs("lastip")=ip
allvalue=int(rs("allvalue"))
info(3)=int(sqr(rs("allvalue")/50))
ljjh_value=rs("allvalue")
ljjh_mvalue=rs("mvalue")
rs.Update
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5">
<title>�s�I</title>
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
<TITLE>�s�I</TITLE>
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
<div align="center">
<tr valign="top">
<td> 
<p align="center"> <%=info(0)%><br>
�� �e �� �A<br><font color=ff0000 size=2>�j�U�w�I��2��</font></p>
<%if info(7)<>"" and info(7)<>"�L" then%>
<%if sfwg=2 then%>
�v��:<font color=red><%=info(7)%></font>�b���Z�\�B���O�j�W�I
<%else%>
�v��:<font color=red><%=info(7)%></font>���b�I
<%end if%>
<%end if%>
�ɶ�:<%=addvalue%>����<br>
<%if info(4)>0  then%>
<div align="center">�z�O<%=info(4)%>�ŷ|��<br>�w�I�����`<%=jhhy%>��</div>
<%else%>
<div align="center">�D�|���I���[��</div>
<%end if%><%end if%>
<p align="left"> 
<%
sjjf=(info(3)+1)*(info(3)+1)*50-ljjh_value%>
�޲z�šG<font color=ff0000><b><%=info(2)%></b></font>��<br>
�԰��šG<font color=ff0000><b><%=int(sqr((ljjh_value/50)))%></b></font>��<br>
��n���G<%=ljjh_mvalue%><br>
�ֿn���G<%=ljjh_value%><br>
�ɯŮt�G<font color=ff0000><b>-<%=sjjf%></b></font><br>
�w�I�ơG<font color=ff0000><b><%=int(addvalue*3*jhhy)%></b></font>�I<br>
<br>
�A�@�W�[<br>
�� �O�G<%if sfwg=2 then%><font color=red><%=int(addvalue1*sfwg*jhhy)%></font>�I<br><%else%><%=int(addvalue1*sfwg*jhhy)%>�I<br><%end if%>
�� �l�G<%=addvalue1*100%>��<br>
�| �O�G<%=addvalue1*0.1%>��<br>
�� ���G<%=addvalue1*0.1%>��<br>
�Z �\�G<%if sfwg=2 then%><font color=red><%=int(addvalue1/2*sfwg*jhhy*100)%></font>�I<br><%else%><%=int(addvalue1/2*sfwg*jhhy)%>�I<br><%end if%>
�k �O�G<%=int(addvalue1/2*jhhy*10)%>�I<br><br>
<%
if bl="�L" or bl="" then
bl=""
else
peioujia=int(addvalue1/4)
conn.execute "update �Τ� set ���O=���O+" & addvalue1/2 & ",�Ȩ�=�Ȩ�+" & peioujia & ",�|���O=�|���O+" & peioujia & ",����=����+" & peioujia & ",�Z�\=�Z�\+" & peioujia & " where �m�W='" & bl & "'"
%>
��Q�G<%=bl%><br>
�� �O:<%=int(addvalue1/4)%>�I<br>
�� ��:<%=int(addvalue1/4)%>��<br>
�Z �\:<%=int(addvalue1/4)%>�I<br>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
end if
%>
</td></tr>
</table></TD>
<TD background="./fq/tb-a4.gif" width="57" height="80"></TD>
</TR>
<TR> 
<TD width="20" height="10"><IMG src="./fq/tb-a8.gif" width="20" height="20" border="0"></TD>
<TD background="./fq/tb-a5.gif" width="102" height="10"></TD>
<TD width="57" height="10"><IMG src="./fq/tb-a10.gif" width="20" height="20" border="0"></TD>
</TR></form></TABLE></BODY></HTML><noscript><xml>