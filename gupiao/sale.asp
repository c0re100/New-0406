<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "叫秈册ぱ秈布"
window.close()
</script>
<%Response.End
end if
if info(0)="" then
%>
<script language="vbscript">
  alert "ぃ秈临⊿Τ祅魁"
  location.href = "../index.asp"
</script>
<%
else
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select 蝗ㄢ from ノめ where ﹎='" & uname & "'"
set rs=conn.execute(sql)
nowmoney=rs("蝗ㄢ")
conn.close
set rs=nothing%>
<!--#include file="jhb.asp"-->
<%sid=Request.Form ("sid")
ushare=abs(Request.Form ("ushare"))
if ushare<1000 then%>
<script language="vbscript">
  alert "程芥虫も1000"
  location.href = "index.asp"
</script>
<%else
sql= "select * from 布 where sid="&sid        
set rs=conn.execute(sql) 
if (rs("讽玡基")-rs("秨絃基"))/rs("秨絃基")<=-0.5  then
%>
<script language="vbscript">
  alert "布禴氨ぃ芥"
  location.href = "index.asp"
</script>

<%else
set rs=nothing
dim uname
sql="select * from め where 眀腹='"&uname&"'"
set rs=conn.execute(sql)
username=rs("眀腹")
nowmoney=rs("戈")

set rs_s=conn.execute ("select 秨絃基,瑈硄布,讽玡基,穨 from 布 where sid="&sid)
set rs_u=conn.execute ("select 计,禦基,キА基 from め where sid="&sid&" and  眀腹='"&username&"'")

dsshare=rs_s("瑈硄布")+ushare


addmoney=ushare*rs_s("讽玡基")

if rs_u.eof then 
Response.Redirect ("nserror.html")
else
dushare=rs_u("计")-ushare
if dushare<0 then
Response.Redirect ("userror.html")

elseif dushare=0 then
sprice=rs_s("讽玡基")
if (sprice-rs_s("秨絃基"))/rs_s("秨絃基")>0.5 then 
sprice=rs_s("秨絃基")*1.5
elseif (sprice-rs_s("秨絃基"))/rs_s("秨絃基")<-0.5 then 
sprice=rs_s("秨絃基")*0.5
end if
shou=(rs_s("讽玡基")-rs_u("キА基"))*ushare
conn.execute "update 布 set ら戳="&date()&",ユ秖=ユ秖+"&ushare&",讽玡基="&sprice&", 瑈硄布="&dsshare&" where sid="&sid
conn.execute "delete from め where 眀腹='"&username&"' and sid="&sid
sql="insert into ユ (眀腹,sid,巨,ユ秖,丁,Μ莉,穨) values ('"&username&"','"&sid&"','芥','"&ushare&"','"&date()&"','"&shou&"','"&rs_s("穨")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update ノめ set 蝗ㄢ=蝗ㄢ+("&addmoney&"*0.99) where  ﹎='"&username&"'"
conn.execute sql
conn.close%>
<!--#include file="jhb.asp"--><%
else
sprice=rs_s("讽玡基")
if (sprice-rs_s("秨絃基"))/rs_s("秨絃基")>0.5 then 
sprice=rs_s("秨絃基")*1.5
elseif (sprice-rs_s("秨絃基"))/rs_s("秨絃基")<-0.5 then 
sprice=rs_s("秨絃基")*0.5
end if
shou=(rs_s("讽玡基")-rs_u("キА基"))*ushare
conn.execute "update 布 set ら戳="&date()&",ユ秖=ユ秖+"&ushare&",讽玡基="&sprice&", 瑈硄布="&dsshare&" where sid="&sid
conn.execute "update め set 计="&dushare&" where 眀腹='"&username&"' and sid="&sid
sql="insert into ユ (眀腹,sid,巨,ユ秖,丁,Μ莉,穨) values ('"&username&"','"&sid&"','芥','"&ushare&"','"&date()&"','"&shou&"','"&rs_s("穨")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update ノめ set 蝗ㄢ=蝗ㄢ+("&addmoney&"*0.99) where  ﹎='"&username&"'"
conn.execute sql
conn.close
end if%>
<!--#include file="jhb.asp"--><%
Randomize
sj1=int(10*rnd)+1
if sj1>5 then
sql="update 布 set 讽玡基=讽玡基*"&(1-(sj1-1)/500)&" where sid="&sid
conn.execute sql 
else
sql="update 布 set 讽玡基=讽玡基*"&(1+sj1/500)&" where sid="&sid
conn.execute sql 
end if
end if
%>

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../style1.css">
<title>布ユΘ</title>
<link rel="stylesheet" href="../../CSS/STYLE.CSS">
</head>

<body bgcolor="#FFFEF4">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr bgcolor="#FFFAE1"> 
    <td width="100%" height="20"> 
      <p align="center"><span class="p"><font color="#006633">竒Θ芥布瓳^_^</font></span>
    </td>
  </tr>
</table>

<dl>
  <dd align="center"><p align="center"><a href="index.asp">布ユ芔</a></p>
  </dd>
</dl>
</body>
</html>
<%end if
end if
end if


%>
<%conn.close
set conn=nothing
%>
