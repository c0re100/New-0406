<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
my=info(0)
if my="" then response.redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs=conn.execute("select 體力 from 用戶 where id="&info(9))
tl=rs("體力")
if tl<0 then
  conn.execute("update 用戶 set 狀態='死' where id="&info(9))
  Session.Abandon
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
  response.write "<script language='javascript'> alert('你好像死了啊！！！'); window.close(); </script>"
    response.end
end if
set rs=nothing
points=session("points")
rpg=session("RPG")
if points="" then points=0
a=Request.Querystring("a")
if a="tou" then
  rpg_times=Request.cookies("rpg_times")
  if rpg_times="" then rpg_times=0
  if rpg_times>3 then

	set rs=nothing	
	conn.close
	set conn=nothing  
    response.write "<script language='javascript'> alert('一天隻能玩三次！！！'); window.close(); </script>"
   response.end
  end if
  rpg_times=rpg_times+1
  Randomize timer
  s=rnd/(rnd+1)+1
  r=int(rnd*8)+1
  points=points+r
  if points>70 then
     points=0
     response.cookies("rpg_times")=rpg_times
     response.cookies("rpg_times").expires=now()+1
  end if
  session("points")=points
  session("RPG")=""
  response.write "<script language='javascript'> alert ('你投了"& r &"點'); location.href='index.asp?"& s &"'; </script>"
  response.end
else
  Randomize timer
  r=rnd/(rnd+1)+1
%>
<html>
<head>
<title>靈劍江湖--甩股子</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "verdana"; font-size: 12px}
a {  cursor: help}
-->
</style>
</head>
<body bgcolor="#fffddf" text="#660000">
<table width="100%" border="1" cellspacing="5" cellpadding="3" bordercolordark="#333333" bordercolorlight="#FFFFFF">
  <tr align="center"> 
    <td colspan="3" height="30">在途中出現<img src="h.gif" width="32" height="32">就有事件發生 
      目前地點：<font color=red><%=points%> </font>&nbsp;&nbsp;甩股子<a href="index.asp?a=tou&<%=r%>"><img src="run.gif" width="38" height="36" border="0"></a><br>
      <br>
      注意：隱藏寶劍隻能拿到一把！！！</td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">0.</font>出發點：祝好運！</td>
    <td width="33%"><font color="#000066">25.</font>公園： <%
if points=25 then
 if rpg="" then
      conn.execute("update 用戶 set 銀兩=銀兩+1000 where id="&info(9))
      session("RPG")=true
%> <img src="h.gif" width="32" height="32">發現1000兩銀子！！！ <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">50.</font>公園：<%
if points=50 then
 if rpg="" then
  set rs=conn.execute("select id from 物品 where 物品名='烈火劍' and 擁有者="&info(9))
    if rs.eof and rs.bof then
     conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,sm,數量,銀兩,等級,會員) values('烈火劍',"&info(9)&",'左手',0,0,0,5000,10000,2004114,2004114,1,8000000,120,"&aaao&")")
else
id=rs("id")
 conn.execute("update 物品 set 數量=數量+1 where id="&id)
end if
 rs.close : set rs=nothing
    session("RPG")=true
%><img src="h.gif" width="32" height="32">發現烈火劍！！！ <%
end if
end if%> </td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">5.</font>公園： <%
if points=5 then
if rpg="" then
conn.execute("update 用戶 set 銀兩=銀兩+5000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">發現5000兩銀子！！！ <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">28-29.</font>武館： <%
if points=28 or points=29 then
if rpg="" then
conn.execute("update 用戶 set 內力=內力+500 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">拜師學武內力提升300！！！ <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">54-55.</font>懸崖： <%
if points=54 or points=55 then
if rpg="" then
conn.execute("update 用戶 set 體力=體力-1000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">跌落懸崖傷體力1000！！！ <%
end if
end if%> </td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">9.</font>公共汽車： <%
if points=9 then
if rpg="" then
conn.execute("update 用戶 set 銀兩=銀兩-5000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">丟失5000兩銀子！！！ <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">35.</font>春院： <%
if points=35 then
if rpg="" then
conn.execute("update 用戶 set 法力=法力+2 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">快活了一番，法力提升2！！！ <%
end if
end if%></td>
    <td width="34%"><font color="#000066">60.</font>懸崖： <%
if points=60 then
if rpg="" then
set rs=conn.execute("select id from 物品 where 物品名='相思劍' and 擁有者="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into 物品(物品名,擁有者,類型,防御,體力,內力,堅固度,攻擊,說明,sm,數量,銀兩,等級,會員) values('相思劍',"&info(9)&",'右手',0,0,0,40000,15000,2004113,2004113,1,5000000,140,"&aaao&")")
else
id=rs("id")
 conn.execute("update 物品 set 數量=數量+1 where id="&id)
end if
rs.close : set rs=nothing
session("RPG")=true
%><img src="h.gif" width="32" height="32">跌落懸崖傷體力<br>
      500，不料發現了一把寶劍！！！ <%
end if
end if%></td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">13-14.</font>避雨樹： <%
if points=13 or points=14 then
if rpg="" then
set rs=conn.execute("select id from 物品 where 物品名='雷霆秘籍' and 擁有者="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into 物品(物品名,擁有者,類型,防御,體力,內力,堅固度,攻擊,說明,sm,數量,銀兩,等級,會員) values('雷霆秘籍',"&info(9)&",'物品',0,0,0,40000,15000,99006,99006,1,10000000,0,"&aaao&")")
else
id=rs("id")
conn.execute("update 物品 set 數量=數量+1  where id="& id &" and 擁有者="&info(9))
end if
session("RPG")=true
%><img src="h.gif" width="32" height="32">一陣閃電霹靂後！在一棵古老的大樹上發現一本《雷霆秘籍》! <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">38-39.</font>礦場： <%
if points=38 or points=39 then
if rpg="" then
set rs=conn.execute("select id from 物品 where 物品名='黑石' and 擁有者="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into 物品(物品名,擁有者,類型,攻擊,防御,體力,內力,堅固度,說明,sm,數量,銀兩,等級,會員) values('黑石',"&info(9)&",'物品',0,0,0,0,0,2003305,2003305,1,10000,0,"&aaao&")")
else
id=rs("id")
 conn.execute("update 物品 set 數量=數量+1 where id="&id)
end if
session("RPG")=true
%><img src="h.gif" width="32" height="32">在礦場閑逛撿到一塊黑石！！！ <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">64-65.</font>夜總會： <%
if points=64 or points=65 then
if rpg="" then
conn.execute("update 用戶 set 內力=內力-500 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">黑幫火拼損失內力500！！！ <%
end if
end if%> </td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">19-20.</font>建築工地： <%
if points=19 or points=20 then
if rpg="" then
set rs=conn.execute("select id from 物品 where 物品名='金屬板' and 擁有者="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into 物品(物品名,擁有者,類型,攻擊,防御,體力,內力,堅固度,說明,sm,數量,銀兩,等級,會員) values('金屬板',"&info(9)&",'物品',0,0,0,0,0,2003304,2003304,1,10000,0,"&aaao&")")
else
id=rs("id")
 conn.execute("update 物品 set 數量=數量+1 where id="&id)
end if
session("RPG")=true
%><img src="h.gif" width="32" height="32">參觀建築工作順手盜取一塊優質金屬板！！！ <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">45.</font>音樂會： <%
if points=45 then
if rpg="" then
conn.execute("update 用戶 set 體力=體力+1000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">體力恢復1000！！ <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">70.</font>無人島： <%
if points=70 then
if rpg="" then
conn.execute("update 用戶 set 銀兩=銀兩+8000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">發現8000兩銀子！！！ <%
end if
end if%> </td>
  </tr>
</table>
</body>
</html>
<%
end if
conn.close
set conn=nothing
%>