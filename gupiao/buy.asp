<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "叫秈册ぱ秈布"
window.close()
</script>
<%Response.End
end if
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select 蝗ㄢ from ノめ where ﹎='" & uname & "'"
set rs=conn.execute(sql)
nowmoney=rs("蝗ㄢ")
conn.close
set rs=nothing %>
<!--#include file="jhb.asp"-->
<%sid=Request.Form ("sid")
ushare=abs(Request.Form ("ushare"))
if ushare<1000 then%>
<script language="vbscript">
  alert "程禦虫も1000"
  location.href = "index.asp"
</script>
<%else
sql= "select * from 布 where sid="&sid        
set rs=conn.execute(sql) 
if (rs("讽玡基")-rs("秨絃基"))/rs("秨絃基")>=0.5  then
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
  alert "布害氨ぃ禦"
  location.href = "index.asp"
</script>


<%else
set rs=nothing

dim uname
sql="select * from め where 眀腹='"&uname&"'"
set rs=conn.execute(sql)
username=rs("眀腹")
sql= "select sum(ユ秖) As 羆 from ユ where 眀腹='"&username&"' and 巨='禦'"
set rs_j=conn.execute(sql) 
if  not rs_j.eof then
yi=rs_j("羆")
if yi+ushare>20000000 then 
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
  alert "磷めカ砏﹚床め–ぱ唉潦禦2000窾布ぃ禦硂或"
  location.href = "index.asp"
</script>
<%

elseif ushare>200000000 then 
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
  alert "磷めカ砏﹚床め–ぱ唉潦禦20000窾布芥ぃ计秖"
  location.href = "index.asp"
</script>
<%else
set rs_s=conn.execute("select 瑈硄布,秨絃基,讽玡基,穨 from 布 where sid="&sid)
set rs_u1=conn.execute("select 计,眀腹,禦基 from め where sid="&sid)

dsshare=rs_s("瑈硄布")-ushare

if nowmoney>=ushare*rs_s("讽玡基") and dsshare>=0 then
   delmoney=ushare*rs_s("讽玡基")


set rs_u=conn.execute ("select 计,禦基 from め where sid="&sid&" and 眀腹='"&username&"'")

if rs_u.eof then
sprice=rs_s("讽玡基")
if (sprice-rs_s("秨絃基"))/rs_s("秨絃基")>0.5 then 
sprice=rs_s("秨絃基")*1.5
elseif (sprice-rs_s("秨絃基"))/rs_s("秨絃基")<-0.5 then 
sprice=rs_s("秨絃基")*0.5
end if
sql="update 布 set 讽玡基="&sprice&", 瑈硄布="&dsshare&",ユ秖=ユ秖+"&ushare&" where sid="&sid
conn.execute sql
sql="insert into め (眀腹,sid,禦基,キА基,计,穨,丁) values ('"&username&"','"&sid&"','"&rs_s("讽玡基")&"','"&rs_s("讽玡基")&"','"&ushare&"','"&rs_s("穨")&"','"&date()&"')"
conn.execute sql
sql="insert into ユ (眀腹,sid,巨,ユ秖,丁,穨) values ('"&username&"','"&sid&"','禦','"&ushare&"','"&date()&"','"&rs_s("穨")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update ノめ set 蝗ㄢ=蝗ㄢ-"&delmoney&" where  ﹎='"&username&"'"
conn.execute sql
conn.close
set rs=nothing %>
<!--#include file="jhb.asp"--><%
else
yuan=rs_u("计")
inprice=rs_u1("禦基")
sprice=rs_s("讽玡基")
if (sprice-rs_s("秨絃基"))/rs_s("秨絃基")>0.5 then 
sprice=rs_s("秨絃基")*1.5
elseif (sprice-rs_s("秨絃基"))/rs_s("秨絃基")<-0.5 then 
sprice=rs_s("秨絃基")*0.5
end if
uprice=(rs_s("讽玡基")*ushare+rs_u("禦基")*yuan)/(ushare+yuan)
sql="update 布 set 讽玡基="&sprice&", 瑈硄布="&dsshare&",ユ秖=ユ秖+"&ushare&" where sid="&sid
conn.execute sql
sql="update め set キА基="&uprice&",计="&rs_u("计")&"+"&ushare&" where 眀腹='"&username&"' and sid="&sid
conn.execute sql
sql="insert into ユ (眀腹,sid,巨,ユ秖,丁,穨) values ('"&username&"','"&sid&"','禦','"&ushare&"','"&date()&"','"&rs_s("穨")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update ノめ set 蝗ㄢ=蝗ㄢ-"&delmoney&" where  ﹎='"&username&"'"
conn.execute sql
conn.close
end if%>
<!--#include file="jhb.asp"--><%
Randomize
sj1=int(10*rnd)+1
if sj1>4 then
sql="update 布 set 讽玡基=讽玡基*"&(1-(sj1-1)/500)&" where sid="&sid
conn.execute sql 
else
sql="update 布 set 讽玡基=讽玡基*"&(1+sj1/500)&" where sid="&sid
conn.execute sql 
end if
elseif dsshare<0 then
Response.Redirect ("serror.html")
else 
Response.Redirect ("uerror.html")
end if
end if
%>

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../style1.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function button1_onclick() {
window.history.go(-2);
}

//-->
</script>

<title>布ユΘ</title>
<link rel="stylesheet" href="../../CSS/STYLE.CSS">
</head>

<body bgcolor="#FFFEF4">
<table border="0" width="100%" bgcolor="#FBE651" cellspacing="0" cellpadding="0">
  <tr bgcolor="#FFFAE1"> 
    <td width="100%" height="20"> 
      <p align="center"><span class="p"><font color="#006633">潦禦Θ</font></span>
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











