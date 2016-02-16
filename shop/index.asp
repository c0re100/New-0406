<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
sql="SELECT * FROM 商店"
rs1.open sql,conn1,1,1
if rs1.EOF or rs1.BOF then Response.Redirect "../error.asp?id=484"
%><title>商店一覽</title>


<BODY bgColor=#CCCCFF>
<DIV ALIGN="LEFT"><font color=#ff0000 size=3>商店一覽</font><BR><BR><FONT SIZE="2">請選擇你要購物的商店<BR>
  要申請商店者,請點<a href="creat.htm">這裡申請</a>! <BR>
  </FONT><%
while not rs1.eof 
%></DIV><center> <a HREF="#" ONCLICK="window.open('mai.asp?shopname=<%=rs1("ID")%>','yuanou','scrollbars=yes,resizable=no,width=665,height=510')"> 
<IMG SRC="1.gif" WIDTH="95" HEIGHT="84" BORDER="0" ALT="<%=rs1("商店名")%>" ></a><%=rs1("商店名")%>---老板:<%=rs1("店主")%><BR><BR></center><%
rs1.MoveNext 
wend
set rs1=nothing
conn1.close
set conn1=nothing
%><BR><BR><BR><BR> <center><%if info(2)>10 then %> <a href=delshop.asp>管理</a><%end if%></center>