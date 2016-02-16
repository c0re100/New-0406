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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select count(*) from 用戶 where 介紹人='"&info(0)&"'",conn
jsren=rs(0)
rs.close
rs.open "Select * from 集合 where id=1",conn
gfmoney=rs("gfmoney")
zmmoney=rs("zmmoney")
zlmoney=rs("zlmoney")
tzmoney=rs("tzmoney")
dzmoney=rs("dzmoney")
lznum=rs("lznum")
zsnum=rs("zsnum")
rs.close
set rs=conn.execute ("Select count(*) from 用戶 where 門派='"&info(5)&"'")
renshu=rs(0)-1
rs.close
rs.open "Select 領錢,會員等級,身份,門派,銀兩,存款,等級 from 用戶 where id="&info(9),conn
mdate=rs("領錢")
hydd=rs("會員等級")
shenfen=rs("身份")
gf=rs("門派")
if  DateDiff("d",rs("領錢"),now())=0 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('天呀，["&info(0) &"]今天你來領過錢的！忘記了？');window.close();</script>"
response.end
end if
yin=rs("銀兩")+rs("存款")
if yin>(rs("會員等級")+1)*500000000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('天呀，["&info(0) &"]你現在有這麼多的銀兩你還想要呀！！');window.close();</script>"
response.end
end if
dj=rs("等級")
hy=rs("會員等級")
if trim(shenfen)="掌門" then

	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+zmmoney
	conn.execute("update 用戶 set 銀兩=銀兩+"&money&",領錢=now() where id="&info(9))
	
end if
if trim(shenfen)="長老" then
	'算錢數
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+zlmoney
	conn.execute("update 用戶 set 銀兩=銀兩+"&money&",領錢=now() where id="&info(9))
	
end if
if trim(shenfen)="堂主" then
	'算錢數
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+tzmoney
	conn.execute("update 用戶 set 銀兩=銀兩+"&money&",領錢=now() where id="&info(9))
	
end if
if trim(gf)="官府" then
    '算錢數
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+gfmoney
	conn.execute("update 用戶 set 銀兩=銀兩+"&money&",領錢=now() where id="&info(9))
end if
if trim(shenfen)<>"掌門" and trim(shenfen)<>"長老" and trim(shenfen)<>"堂主" and trim(gf)<>"官府" then
	'算錢數
	money=(hy+1)*20000+cint(renshu)*5000+jsren*1000+dj*1500+dzmoney
	conn.execute("update 用戶 set 銀兩=銀兩+"&money&",領錢=now() where id="&info(9))
end if
rs.close
rs.open "select id from 物品 where 物品名='萬年靈芝' and 擁有者=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,內力,體力,數量,銀兩,說明,會員) values ('萬年靈芝',"&info(9)&",'藥品',0,200000,2,5000000,136,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update 物品 set 數量=數量+"&lznum&",會員="&aaao&" where id="& id &" and 擁有者="&info(9)
				
		        end if
if hy>0 then
rs.close
	rs.open "select id from 物品 where 物品名='魔力鑽石' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,內力,體力,數量,銀兩,會員) values ('魔力鑽石',"&info(9)&",'法器',0,0,2,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 銀兩=200000,數量=數量+"&zsnum&",會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
end if
rs.close
rs.open "Select 門派,身份,等級 from 用戶 where id="&info(9),conn
%>
<html>
<head>
<title>工資領取處</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<STYLE type=text/css>
TD {FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt}
BODY {FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt}
SELECT {FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</STYLE>
</head>
<body bgcolor="#0099FF" text="#FFFFFF" leftmargin="0">
<div align="center">
<p><font color="#000000"><b><font size="+3">工資領取處</font></b></font> <br>
<br>
<font size=+1> <b><%=info(0)%></b></font>你是<font color="#0000FF">[<%=rs("門派")%>]</font>
的<font color="#FF0000"> [<%=rs("身份")%>] </font>拉人：<font color="#FF0000"><b><%=jsren%>個</b></font>
戰鬥等級:<font color="#FF0000"><%=rs("等級")%>級</font>
<%if trim(rs("身份"))="掌門" then%>
你有弟子:<font color="#FF0000"><b><%=renshu%>個</b></font> <%end if%><br>多多努力，多多拉人！
<br>
今天你從江湖領取到了<b><font color="#FF0000"><%=money%>兩</font></b>，小心保存，不要亂花！
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
<br>
</p>
  <p align="center"><br>
    <font color="#FF0000" size="+1"><b>工資計算方法</b></font><br>
    <font color="#0000FF">官府：</font>會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500+<%=gfmoney%> 
    <br>
    <font color="#0000FF">掌門：</font>會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500+<%=zmmoney%><br>
    <font color="#0000FF">長老：</font>會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500+<%=zlmoney%><br>
    <font color="#0000FF">堂主：</font>會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500+<%=tzmoney%> 
    <br>
    <font color="#0000FF">弟子：</font>會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500+<%=dzmoney%> 
    <br>
    <font color="#0000FF">會員：</font>會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500+<%=dzmoney%> 
    <br>
    <br>
    <font color="#99FFCC">非會員每天可以在此領取<%=lznum%>個萬年靈芝<br>
    會員每天可以在此領取<%=zsnum%>個魔力鑽石</font></p> </p>
  <p align="center">&nbsp;</p>
<p align="center">
<input type=button value=關閉窗口 onClick='window.close()' name="button">
</p>
</div>
</body>
</html>
