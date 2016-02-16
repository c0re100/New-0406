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
rs.open "select 非會員泡點,會員泡點1,會員泡點2,會員泡點3,會員泡點4,會員泡點5,會員泡點6,會員泡點7,會員泡點8,會員泡點9,會員泡點10 FROM 文本 ",conn
pao0=rs("非會員泡點")
pao1=rs("會員泡點1")
pao2=rs("會員泡點2")
pao3=rs("會員泡點3")
pao4=rs("會員泡點4")
pao5=rs("會員泡點5")
pao6=rs("會員泡點6")
pao7=rs("會員泡點7")
pao8=rs("會員泡點8")
pao9=rs("會員泡點9")
pao10=rs("會員泡點10")
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
if info(7)<>"" and info(7)<>"無" then
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
sql="SELECT 內力,銀兩,會員費,金幣,武功,殺人數,操作時間,法力,配偶,grade,allvalue,泡豆點數,mvalue,lasttime,lastip FROM 用戶 WHERE id="&info(9)
rs.open sql,conn,1,3
killsj=DateDiff("n",rs("操作時間"),now())
if killsj>30 and rs("殺人數")>=10 then
conn.execute "update 用戶 set 殺人數=殺人數-1,操作時間=now() where id="&info(9)
end if

if Not(rs.Eof and rs.Bof) then
rs("銀兩")=rs("銀兩")+int(addvalue1*jhhy*100)
rs("會員費")=rs("會員費")+int(addvalue1*jhhy*0.1)
rs("金幣")=rs("金幣")+int(addvalue1*jhhy*0.1)
rs("內力")=rs("內力")+int(addvalue1*sfwg*jhhy*10)
rs("武功")=rs("武功")+int(addvalue1/2*sfwg*jhhy*2)
rs("法力")=rs("法力")+int(addvalue1/2*jhhy*2)
bl=rs("配偶")
grade=rs("grade")
rs("allvalue")=rs("allvalue")+int(addvalue*jhhy*9)
rs("泡豆點數")=rs("泡豆點數")+addvalue
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
<title>存點</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
<HTML>
<HEAD>
<TITLE>存點</TITLE>
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
目 前 狀 態<br><font color=ff0000 size=2>大廳泡點為2倍</font></p>
<%if info(7)<>"" and info(7)<>"無" then%>
<%if sfwg=2 then%>
師傅:<font color=red><%=info(7)%></font>在場武功、內力大增！
<%else%>
師傅:<font color=red><%=info(7)%></font>不在！
<%end if%>
<%end if%>
時間:<%=addvalue%>分鐘<br>
<%if info(4)>0  then%>
<div align="center">您是<%=info(4)%>級會員<br>泡點為正常<%=jhhy%>倍</div>
<%else%>
<div align="center">非會員點不加倍</div>
<%end if%><%end if%>
<p align="left"> 
<%
sjjf=(info(3)+1)*(info(3)+1)*50-ljjh_value%>
管理級：<font color=ff0000><b><%=info(2)%></b></font>級<br>
戰鬥級：<font color=ff0000><b><%=int(sqr((ljjh_value/50)))%></b></font>級<br>
月積分：<%=ljjh_mvalue%><br>
累積分：<%=ljjh_value%><br>
升級差：<font color=ff0000><b>-<%=sjjf%></b></font><br>
泡點數：<font color=ff0000><b><%=int(addvalue*3*jhhy)%></b></font>點<br>
<br>
你共增加<br>
內 力：<%if sfwg=2 then%><font color=red><%=int(addvalue1*sfwg*jhhy)%></font>點<br><%else%><%=int(addvalue1*sfwg*jhhy)%>點<br><%end if%>
銀 子：<%=addvalue1*100%>兩<br>
會 費：<%=addvalue1*0.1%>元<br>
金 幣：<%=addvalue1*0.1%>個<br>
武 功：<%if sfwg=2 then%><font color=red><%=int(addvalue1/2*sfwg*jhhy*100)%></font>點<br><%else%><%=int(addvalue1/2*sfwg*jhhy)%>點<br><%end if%>
法 力：<%=int(addvalue1/2*jhhy*10)%>點<br><br>
<%
if bl="無" or bl="" then
bl=""
else
peioujia=int(addvalue1/4)
conn.execute "update 用戶 set 內力=內力+" & addvalue1/2 & ",銀兩=銀兩+" & peioujia & ",會員費=會員費+" & peioujia & ",金幣=金幣+" & peioujia & ",武功=武功+" & peioujia & " where 姓名='" & bl & "'"
%>
伴侶：<%=bl%><br>
內 力:<%=int(addvalue1/4)%>點<br>
銀 兩:<%=int(addvalue1/4)%>兩<br>
武 功:<%=int(addvalue1/4)%>點<br>
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