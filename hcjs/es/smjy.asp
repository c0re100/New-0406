<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();}</script>"
	Response.End 
end if

%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
myname=trim(request.querystring("myname"))
PageSize = 15
rs.open "delete * from 交易市場 where 時間<now()-10 and 方式<>'保險櫃'",conn,3,3
if myname="" then
	rs.open "Select id,物品名,擁有者,數量,對像,時間,說明,銀兩,二手價,對像 From 交易市場 where 方式='交易' Order by 時間 DESC",conn,3,3
else
	rs.open "Select id,物品名,擁有者,數量,對像,時間,說明,銀兩,二手價,對像 From 交易市場 where 方式='交易' and (對像='"& myname &"' or 擁有者="& info(9) &") Order by 時間 DESC",conn,3,3
end if

rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>商貿交易</title>
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="0" height="24" width="712" cellspacing="0" cellpadding="0"
align="center">
  <tbody> 
  <tr>
    <td height="15" width="79%" bgcolor="#99CCFF"><font color="#669966"> <font color="#FFFFFF"><b><font color="#0000FF">商貿交易</font>(</b><font color="#000000">注意10天交易不成功的過期物品我們將送到垃圾場處理!</font><b>）<a                    href="smjy.asp?myname=<%=info(0)%>"><font color="#990000">查找我的交易</font></a> 
      </b></font></font><font face="幼圓"><a href="javascript:this.location.reload()"><font color="#990000">刷新此頁面</font></a></font><font color="#990000"><b> 
      </b></font></td>
    <td height="15" width="21%" bgcolor="#99CCFF"> 
      <div align="right"><font color="#669966"><font color="#FFFFFF"><b><a href="wupin.asp" target="_blank"><font color="#990000">處理自己物品點這裡</font></a></b></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="712" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
<td width="100%">
        <table border="1" cellspacing="1" cellpadding="0" width="712"
bordercolorlight="#EFEFEF" align="center">
          <tr bgcolor="#FFFFFF">
            <td width="11%" height="16"> 
              <div align="center"><font color="#FF6600"> 物品名</font></div>
</td>
            <td width="8%" height="16"> 
              <div align="center"><font color="#FF6600"> 擁有者</font></div>
</td>
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600"> 數量</font></div>
</td>
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600">對像</font></div>
</td>
            <td width="13%" height="16"> 
              <div align="center"><font color="#FF6600"> 時 間 </font></div>
</td>
            <td width="23%" height="16"> 
              <div align="center"><font color="#FF6600">說 明</font></div>
</td>
            <td width="7%" height="16"> 
              <div align="center"><font color="#FF6600"> 原價</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600">現價</font></div>
</td>
            <td width="13%" height="16"> 
              <div align="center"><font color="#FF6600">購買</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr>
            <td width="11%" height="21"> 
              <div align="center"> <font color="#FFFFCC"><%=rs("物品名")%></font>
</div>
</td>
            <td width="8%" height="21"> 
              <div align="center"> <%=rs("擁有者")%> </div>
</td>
            <td width="9%" height="21"> 
              <div align="center"> <%=rs("數量")%> </div>
</td>
            <td width="9%" height="21"> 
              <div align="center"><%=rs("對像")%></div>
</td>
            <td width="13%" height="21"> 
              <div align="center"> <%=rs("時間")%> </div>
</td>
            <td width="23%" height="21"> 
              <div align="left"></div>
<%=rs("說明")%></td>
            <td width="7%" height="21"> 
              <div align="center"> <%=rs("銀兩")%> </div>
</td>
<td width="7%" height="21">
<div align="center"> <%=rs("二手價")%> </div>
</td>
            <td width="13%" height="21"> 
              <div align="center"> 
                <% if info(2)>=9 then%>
                <a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">刪除</font></a> 
                <%end if%>
                <%if rs("對像")=info(0) then%>
                <a href="smjy1.asp?id=<%=rs("id")%>"><b><font color="#FFFFFF">要了</font></b></a> 
                <%else%>
                <font color="#FFFFFF">no交易</font> 
                <%end if%>
              </div>
            </td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
        <table border="0" cellspacing="1" cellpadding="1" width="712" bordercolorlight="#EFEFEF" align="center">
          <tr>
<td align="left" width="37%" height="2">[共<font color="red"><b><%=pa%></b></font>頁][<font
color="red"><b><%=mu%></b></font>次交易]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="smjy.asp?page=<%=page-1%>">上一頁</a>][第<%=page%>頁][<a                    href="smjy.asp?page=<%=page+1%>">下一頁</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 數量 from 交易市場 where 方式='交易'")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>

</body>