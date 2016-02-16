<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")


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
if sheepname="" or len(sheepname)>16 then
response.redirect"warning.htm"
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open"select 銀兩 from 用戶 where 姓名='"&info(0)&"'",conn,1,1
if rs("銀兩")<100000 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"您沒有足夠的錢！",0,"FLUSH"
history.back
</script>
<%
else
money=rs("銀兩")-100000
conn.execute("update 用戶 set 銀兩='"&money&"'  where 姓名='"&info(0)&"'")
rs.close%>
<!--#include file="data1.asp"-->

<%rs.open"select user from sheep where user='"&info(0)&"'",conn,1,1
if not rs.bof then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"一個人隻能領養一個小孩",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select user from sheep",conn,1,1
if rs.recordcount>10000 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"城市裡面的羊已經超過了10000隻，你不能再領養",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select cleaninit,happyinit,healthinit,milkinit,lifeinit,hungryinit from rules",conn,1,1
cleaninit=rs("cleaninit")
sheephappyinit=rs("happyinit")
sheephealthinit=rs("healthinit")
milkinit=rs("milkinit")
lifeinit=rs("lifeinit")
hungryinit=rs("hungryinit")
sheepdate=date
feeddate=DATEADD("d",-1,DATE)
workload=0
rs.close
conn.execute"insert into sheep(sheephappy,sheephealth,user,sheepname,sheepdate,life,milk,hungry,workload,clean,feeddate,feedsheepday,logintoday) values('"&sheephappyinit&"','"&sheephealthinit&"','"&info(0)&"','"&sheepname&"','"&sheepdate&"','"&lifeinit&"','"&milkinit&"','"&hungryinit&"','"&workload&"','"&cleaninit&"','"&feeddate&"','0','"&feeddate&"')"
rs.open"select id from sheep",conn,1,1
rs.movelast
sheepid=rs("id") '小羊編號
conn.execute"update sheep set user='"&info(0)&"' where sheepname='"&sheepname&"'"

%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

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
        <p align="center"><font size="2">恭喜！您已經成功領養了<%=sheepname%>，您領養的孩子的編號是：</font><%=rs("id")%></td>
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
conn.close 
end if 
end if
end if 
end if %>
