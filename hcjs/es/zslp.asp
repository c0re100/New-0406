<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
myname=trim(request.querystring("myname"))
PageSize = 15
rs.open "delete * from 交易市場 where 時間<now()-5 and 方式<>'保險櫃'",conn,3,3
if myname="" then
rs.open "Select id,物品名,擁有者,數量,對像,時間,說明,銀兩,對像 From 交易市場 where 方式='贈送' Order by 時間 DESC",conn,3,3
else
rs.open "Select id,物品名,擁有者,數量,對像,時間,說明,銀兩,對像 From 交易市場 where 方式='贈送' and  (對像='"& myname &"' or 擁有者="& info(9)&") Order by 時間 DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>情意禮品店</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body leftmargin="0" topmargin="0" background="../../ff.jpg">
<table border="0" height="24" width="700" cellspacing="0" cellpadding="0"
align="center">
  <tbody>
    <tr bgcolor="#FF0000"> 
      <td height="15" width="79%"><font color="#669966"> <font color="#FFFFFF"><font color="#000000">注意10天交易不成功的過期物品我們將送到垃圾場處理!</font><b><a                    href="zslp.asp?myname=<%=info(0)%>"> 
        </a></b><a href="zslp.asp?myname=<%=info(0)%>">查找我的禮品</a> 
        </font><font color="#669966"><font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()">刷新此頁面</a></font></font></font></td>
      <td height="15" width="21%"> 
        <div align="right"><font color="#669966"><font color="#FFFFFF"><a href="wupin.asp" target="_blank">處理自己物品點這裡</a></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="700" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
<td width="100%">
<table border="0" cellspacing="2" cellpadding="2" width="700"
bordercolorlight="#EFEFEF">
          <tr bgcolor="#FFFFFF">
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> 物品名</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> 送禮人</font></div>
</td>
<td width="5%" height="16">
<div align="center"><font color="#FF6600"> 數量</font></div>
</td>
<td width="8%" height="16">
<div align="center"><font color="#FF6600">對像</font></div>
</td>
<td width="17%" height="16">
<div align="center"><font color="#FF6600"> 時 間 </font></div>
</td>
<td width="29%" height="16">
<div align="center"><font color="#FF6600">說 明</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> 原值</font></div>
</td>
<td height="16" colspan="2">
<div align="center"><font color="#FF6600">操作</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr>
<td width="7%" height="29">
<div align="center"> <font color="#0000FF"><%=rs("物品名")%></font>
</div>
</td>
<td width="7%" height="29">
<div align="center"> <%=rs("擁有者")%> </div>
</td>
<td width="5%" height="29">
<div align="center"> <%=rs("數量")%> </div>
</td>
<td width="8%" height="29">
<div align="center"><font color="#0000FF"><%=rs("對像")%></font></div>
</td>
<td width="17%" height="29">
<div align="center"> <%=rs("時間")%> </div>
</td>
<td width="29%" height="29">
<div align="left"></div>
<%=rs("說明")%></td>
<td width="7%" height="29">
<div align="center"> <%=rs("銀兩")%> </div>
</td>
<td width="12%" height="29">
<div align="center"><% if info(2)>=9 then%><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">刪除</font></a><%end if%>
<%if len(rs("說明"))>5 then zy=left(rs("說明"),4)
if rs("對像")=info(0) and zy<>"自作多情" then%>
<a href="zslp1.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>謝謝了</b></font></a>
<%else%>
<font color="#0000FF"><b>no我的</b></font>
<%end if%></div>
</td>
<td width="8%" height="29">
<%if len(rs("說明"))>5 then zy=left(rs("說明"),4)

if rs("對像")=info(0) and zy<>"自作多情" then%>
<div align="center"><a href="lpjj.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>自作多情</b></font></a></div>
<%else%>
<div align="center"><font color="#0000FF"><b>no我的</b></font></div>
<%end if%>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
conn.close
set rs=nothing
set conn=nothing

%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>頁][<font
color="red"><b><%=mu%></b></font>次賺送]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="zslp.asp?page=<%=page-1%>">上一頁</a>][第<%=page%>頁][<a                    href="zslp.asp?page=<%=page+1%>">下一頁</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 數量 from 交易市場 where 方式='贈送'")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>
</body>
