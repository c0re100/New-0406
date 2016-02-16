<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
sql="SELECT 商店名,店主,注冊資金,id FROM 商店 where 店主='"&info(0)&"'"
rs1.open sql,conn1,1,1
if rs1.EOF or rs1.BOF then Response.Redirect "../error.asp?id=484"
dim shopname
shopname=rs1("商店名")
name=rs1("店主")
shopmoney=rs1("注冊資金")
id=rs1("id")
if name<>info(0) then
set rs1=nothing
conn1.close
set conn1=nothing
%>
<script language="vbscript">
alert("這家商店不是你開的！")
history.back(1)
</script>
<%Response.End 
end if
%>
<head>
<title><%=shopname%></title>
</head>

<body bgcolor="#0099CC" BACKGROUND="../linjianww/0064.jpg" text="#000000" topmargin="0">
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="color: #E0E4E0; font-size: 10pt">
  <tr bgcolor="#0033CC"> 
    <td colspan="9" align="center"> <font size="+3" color="#FFFFFF"> 
      <%rs1.close
	  sql="select * from 物品 where 擁有者="&id&" and 類型<>'卡片'  and 裝備=false and 數量>0"
rs.Open sql,conn,1,1
if not rs.EOF then
%>
      物品表</font> </td>
  </tr>
  <tr bgcolor="#0099FF"> 
    <td width="21%" align="center"><font color="#FFFFFF">商品名</font></td>
    <td width="12%" align="center"><font color="#FFFFFF">類型</font></td>
    <td width="10%" align="center"><font color="#FFFFFF">價格</font></td>
    <td width="11%" align="center"><font color="#FFFFFF">數量</font></td>
    <td width="46%" align="center"><font color="#FFFFFF">設 置</font></td>
  </tr>
  <%while not rs.EOF or rs.BOF %>
  <tr> 
    <td width="21%" align="center" background="bg.gif"><font color="#000000"><%=rs("物品名")%></font></td>
    <td width="12%" align="center" background="bg.gif"><font color="#000000"><%=rs("類型")%></font></td>
    <td width="10%" align="center" background="bg.gif"><font color="#000000"><%=rs("銀兩")%></font></td>
    <td width="11%" align="center" background="bg.gif"><font color="#000000"><%=rs("數量")%></font></td>
    <td width="46%" align="center" background="bg.gif"> 
      <form method="POST" action="modify.asp">
        <font color="#000000"> 
        <input type="hidden" name="name" value=<%=rs("物品名")%>>
        <input type="hidden" name=aaa value=<%=rs("擁有者")%>>
        <input type="submit" value="加入商店" name="submit">
        </font> 
      </form>
    </td>
  </tr>
  <% 
rs.MoveNext 
wend 
end if
%>
</table>
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="color: #E0E4E0; font-size: 10pt">
  <tr bgcolor="#0033CC"> 
    <td colspan="9" align="center"> <font color="#FFFFFF" size="+3">商店物品表</font></td>
  </tr>
  <tr bgcolor="#0099FF"> 
    <td width="21%" align="center"><font color="#FFFFFF">商品名</font></td>
    <td width="12%" align="center"><font color="#FFFFFF">類型</font></td>
    <td width="10%" align="center"><font color="#FFFFFF">價格</font></td>
    <td width="11%" align="center"><font color="#FFFFFF">數量</font></td>
    <td width="46%" align="center"><font color="#FFFFFF">調價</font></td>
  </tr>
  <%rs.close
  set rs=nothing
conn.close
set conn=nothing
sql="select * from 商店物品 where 擁有者="&id&" and 裝備=false and 數量>0"
rs1.Open sql,conn1,1,1
if not rs1.EOF then

  while not rs1.EOF or rs1.BOF %>
  <tr> 
    <td width="21%" align="center" background="bg.gif"><font color="#000000"><%=rs1("物品名")%></font></td>
    <td width="12%" align="center" background="bg.gif"><font color="#000000"><%=rs1("類型")%></font></td>
    <td width="10%" align="center" background="bg.gif"><font color="#000000"><%=rs1("銀兩")%></font></td>
    <td width="11%" align="center" background="bg.gif"><font color="#000000"><%=rs1("數量")%></font></td>
    <td width="46%" align="center" background="bg.gif"> 
      <form method="POST" action="modifyshang.asp">
        <font color="#000000"> 
        <input type="hidden" name="name2" value=<%=rs1("物品名")%>>
        <input type="hidden" name=aaa2 value=<%=rs1("擁有者")%>>
        <select name="select" style="border: 1px solid; background-color:#EEFFEE;font-size: 9pt; border-color:
#000000 solid" >
          <option value="a" selected>乘以2</option>
          <option value="b">乘以4</option>
          <option value="c">乘以6</option>
          <option value="d">乘以8</option>
          <option value="e">乘以10</option>
          <option value="f">乘以50</option>
          <option value="g">乘以100</option>
          <option>---------------</option>
          <option value="h">除以2</option>
          <option value="i">除以4</option>
          <option value="j">除以6</option>
          <option value="k">除以8</option>
          <option value="l">除以10</option>
          <option value="m">除以50</option>
          <option value="n">除以100</option>
        </select>
        <input type="submit" value="O K" name="submit2">
        </font> 
      </form>
    </td>
  </tr>
  <% 
rs1.MoveNext 
wend 
set rs1=nothing
conn1.close
set conn1=nothing

end if
%>
</table>
<p align="center"><A HREF="modifyshop.asp">返回</A> <BR>
   