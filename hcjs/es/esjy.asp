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
	rs.open "Select id,物品名,擁有者,數量,時間,說明,銀兩,二手價 From 交易市場 where 方式='二手'  Order by 時間 DESC",conn,3,3
else
	rs.open "Select id,物品名,擁有者,數量,時間,說明,銀兩,二手價 From 交易市場 where 方式='二手'  and  擁有者="& info(9) &" Order by 時間 DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>二手交易</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#660000" background="../../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<table border="0" height="24" width="713" cellspacing="0" cellpadding="0"
align="center" bgcolor="#FF0000">
  <tbody> 
  <tr>
      <td height="15" width="80%"><font color="#669966"> <font color="#FFFFFF"><b>(</b><font color="#FFFFCC">注意10天交易不成功的過期物品我們將送到垃圾場處理!</font><b>）</b></font><font color="#669966"><font color="#FFFFFF"><a  href="esjy.asp?myname=<%=info(0)%>">查找我的二手物品</a></font><font color="#669966"><font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()"> 
        </a></font></font></font></font></td>
      <td width="20%" height="15" bgcolor="#FF0000"> 
        <div align="right"><font color="#669966"><font color="#FFFFFF"><a href="wupin.asp" target="_blank">處理自己物品點這裡</a></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="713" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
      <td width="100%"> 
        <table border="0" cellspacing="2" cellpadding="2" width="713"
bordercolorlight="#EFEFEF" align="center">
          <tr bgcolor="#FF0000"> 
            <td width="10%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                物 品 名</font></div></td>
            <td width="7%" height="11"> 
              <div align="center"><font color="#FFFFFF">擁有者 
                </font></div></td>
            <td width="7%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                數量</font></div></td>
            <td width="16%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                時 間 </font></div></td>
            <td width="25%" height="11"> 
              <div align="center"><font color="#FFFFFF">說 
                明</font></div></td>
            <td width="8%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                原始價錢</font></div></td>
            <td width="8%" height="11"> 
              <div align="center"><font color="#FFFFFF">二手價錢</font></div></td>
            <td width="19%" height="11"> 
              <div align="center"><font color="#FFFFFF">購買</font></div></td>
          </tr>
          <%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
          <tr bgcolor="#006699"> 
            <td width="10%" height="7"> 
              <div align="center"> <font color="#FFFFCC"><%=rs("物品名")%></font> </div></td>
            <td width="7%" height="7"> 
              <div align="center"> <%=rs("擁有者")%> </div></td>
            <td width="7%" height="7"> 
              <div align="center"> <%=rs("數量")%> </div></td>
            <td width="16%" height="7"> 
              <div align="center"> <%=rs("時間")%> </div></td>
            <td width="25%" height="7"> 
              <div align="left"></div>
              <%=rs("說明")%></td>
            <td width="8%" height="7"> 
              <div align="center"> <%=rs("銀兩")%> </div></td>
            <td width="8%" height="7"> 
              <div align="center"> <%=rs("二手價")%> </div></td>
            <td width="19%" height="7"> 
              <div align="center"><b> 
                <% if info(2)>=9 then%>
                </b><a href="del.asp?id=<%=rs("id")%>"><font color="#FFFF00">刪除</font></a><b> 
                <%end if%>
                </b><a href="esjy1.asp?id=<%=rs("id")%>"><font color="#FFFF00">要了</font></a><b> 
                </b></div></td>
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
        <table border="0" cellspacing="1" cellpadding="1" width="713" bordercolorlight="#EFEFEF" align="center">
          <tr>
            <td align="left" width="37%" height="2"><font color="#FFCC00">[共</font><font color="red"><b><%=pa%></b></font><font color="#FFCC00">頁][<b><%=mu%></b>次交易]</font></td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="esjy.asp?page=<%=page-1%>">上一頁</a>][第<%=page%>頁][<a href="esjy.asp?page=<%=page+1%>">下一頁</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 數量 from 交易市場 where 方式='二手'")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>
</body>