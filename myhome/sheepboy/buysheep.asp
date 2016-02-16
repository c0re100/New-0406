<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
nickname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 配偶,性別,銀兩 from 用戶 where id="&info(9),conn
peier=rs("配偶")
if peier="" or peier="無" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你還沒有配偶啊！');location.href = 'indexsheep.asp';}</script>"
Response.End
end if
if rs("性別")="男" then 
  sheep00a1=nickname
  sheep00a2=peiper
  else
   sheep00a1=peiper
   sheep00a2=nickname
end if
if rs("銀兩")>10000000 then 
  sheepxb="男"
 else
  sheepxb="女"
end if


function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function

function cctime (stime)
if len(hour(stime))=1 then
h="0"&hour(stime)
else
h=hour(stime)
end if
if len(minute(stime))=1 then
m="0"&minute(stime)
else
m=minute(stime)
end if
if len(second(stime))=1 then
s="0"&second(stime)
else
s=second(stime)
end if
cctime=h&":"&m&":"&s
end function
'=====================================================

sheepname=request.form("sheepname")
sheepname2=request.form("sheepname2")
if sheepname="" or len(sheepname)>16 or sheepname2="" or len(sheepname)>12 then
response.redirect"warning.htm"
else
rs.close
rs.open"select 銀兩 from 用戶 where 姓名='"&info(0)&"'",conn,1,1
if rs("銀兩")<10 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你沒錢啊！');location.href = 'indexsheep.asp';}</script>"
Response.End
else
money=rs("銀兩")-10
'conn.execute("update 用戶 set 銀兩='"&money&"'  where 姓名='"&info(0)&"'")
rs.close%>
<!--#include file="data1.asp"-->

<%rs.open"select user from sheep where user='"&info(0)&"'",conn,1,1
if not rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('計劃生育暫時隻能生一個孩子！');location.href = 'indexsheep.asp';}</script>"
Response.End
else
rs.close
rs.open"select user from sheep",conn,1,1
if rs.recordcount>10000 then
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('該江湖育兒院裡已經有超過了10000小孩了，你們要想生小孩的話，跟站長聯繫吧，呵呵');location.href = 'indexsheep.asp';}</script>"
Response.End
%>

<!--#include file="data1.asp"-->
<%
else
rs.close
rs.open"select cleaninit,happyinit,healthinit,milkinit,lifeinit,hungryinit,wugonginit,neiliinit,gongjiinit from rules",conn,1,1
cleaninit=rs("cleaninit")
sheephappyinit=rs("happyinit")
sheephealthinit=rs("healthinit")
milkinit=rs("milkinit")
lifeinit=rs("lifeinit")
hungryinit=rs("hungryinit")
wugonginit=rs("wugonginit")
neiliinit=rs("neiliinit")
gongjiinit=rs("gongjiinit")
sheepdate=date
feeddate=DATEADD("d",-1,DATE)
workload=0
rs.close
conn.execute"insert into sheep(sheephappy,sheephealth,sheep001,sheepname,sheepdate,life,milk,wugong,leinl,gongji,hungry,workload,clean,feeddate,feedsheepday,logintoday) values('"&sheephappyinit&"','"&sheephealthinit&"','"&sheep00a1&"','"&sheepname&"','"&sheepdate&"','"&lifeinit&"','"&milkinit&"','"&wugonginit&"','"&neiliinit&"','"&gongjiinit&"','"&hungryinit&"','"&workload&"','"&cleaninit&"','"&feeddate&"','0','"&feeddate&"')"
rs.open"select id from sheep",conn,1,1
rs.movelast
sheepid=rs("id") '小羊編號
conn.execute"update sheep set user='"&info(0)&"' where sheepname='"&sheepname&"'"
conn.execute"update sheep set sheep002='"&sheepname2&"' where sheepname='"&sheepname&"'"
conn.execute"update sheep set sheepxb='"&sheepxb&"' where sheepname='"&sheepname&"'"
 Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

rs.open "select id from 用戶 where id="&info(9),conn
conn.execute("update 用戶 set 銀兩=銀兩-1000000,小孩='有',孩名='"&sheepname&"'  where 姓名='"&info(0)&"'")
conn.execute("update 用戶 set 銀兩=銀兩-1000000,小孩='有',孩名='"&sheepname&"'  where 姓名='"&peier&"'")
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" background="../../images/8.jpg">
<div align="center">
</div>
<div align="center">
  <center>
  <table border="0" width="470" bordercolor="#FFFFFF" cellspacing="1"
  cellpadding="0" height="20">
    <tr>
        <td class="p6"><font size="2">□-您當前的位置---<a
        href="indexsheep.asp">孤兒院</a>-<a
        href="feedsheep.asp">托兒所</a></font></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
  <table border="0" width="470" cellspacing="1" cellpadding="0" height="50">
    <tr>
        <td class="p2" width="100%"> 
          <%if sheepxb="男" then
		  msg="<font color=FF0000>【生小孩】</font>恭喜:"&info(0)&"和"&peier&"成功的生下一個<img src='../Myhome/SHEEPBOY/baby3.gif'>孩子的名字是"&sheepname&"...."

		 		   %>
		            <p align="center"><font size="2">恭喜！你們成功的生下一個<img src="baby3.gif" width="144" height="133">，你們孩子的江湖編號是：</font><%=rs("id")%> 
            <%
  else
   msg="<font color=FF0000>【生小孩】</font>恭喜:"&info(0)&"和"&peier&"成功的生下一個<img src='../Myhome/SHEEPBOY/girl15.gif'>孩子的名字是"&sheepname&"...."

  %>
          <p align="center"><font size="2">恭喜！你們成功的生下一個<img src="girl15.gif" width="45" height="57">，你們孩子的江湖編號是：</font><%=rs("id")%> 
            <%end if%>
          <p align="center">&nbsp;
        </td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <table border="0" width="470" cellspacing="1" cellpadding="0">
    <tr>
      <td class="p3" width="100%">
        <p align="right"><font size="2">快回托兒所開始照顧<%=sheepname%>吧！</a></font></td>
    </tr>
  </table>
</div>
<div align="center">
  <center>
    <font color="#FFFFFF"><a href="javascript:window.close()">關閉窗口</a></font>
</center>
</div>

</body>

</html>
<% 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
end if 
end if 
end if 
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
sd(194)="消息"
sd(195)="大家"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>