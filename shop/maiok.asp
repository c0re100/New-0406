<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("shop.asa")
Conn.Open connstr
shang=Request.QueryString("shang")
leixin=Request.QueryString("lei")
sql="SELECT * FROM 商店 where 商店名='"&shang&"'"
rs.open sql,conn,1,1
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=484"
shopname=rs("ID")
user=rs("店主")

rs.close
rs.Open "SELECT * FROM 商店物品  where 擁有者="&shopname&" and 數量>0",conn
if rs.EOF or rs.BOF then
rs.close
set rs=nothing
conn.close
set conn=nothing
if user=info(0) then
%><head>
<script language="vbscript">
alert("此店暫時沒有物品出售！")
location.href = "modifyshop.asp"
</script>
<%else
%><head>
<script language="vbscript">
alert("此店暫時沒有物品出售！")
location.href = "javascript:window.close()"
</script>
<%
Response.End 
end if
end if%>
<title><%=shang%>---<%=leixin%></title>
</head>

<body bgcolor="#0066CC" LINK="#99FFCC" >
<p align="center"><FONT SIZE="+3" COLOR="#FFFFFF" ><%=shang%></FONT> <A HREF="modifyshop.asp" TARGET="_blank"><FONT SIZE="3" COLOR="#000000"> 
  </FONT></A><FONT SIZE="3" COLOR="#000000"> </FONT> 
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="color: #E0E4E0; font-size: 10pt" background="../linjianww/0064.jpg">
  <tr bgcolor="#0099FF"> 
    <td width="60" align="center" ><font color="#FFFFFF">商品名</font></td>
    <td width="177" align="center" ><font color="#FFFFFF">商品圖片</font></td>
    <td width="266" align="center" ><font color="#FFFFFF">增加值</font></td>
    <td width="67" align="center" ><font color="#FFFFFF">數量</font></td>
    <td width="98" align="center" ><font color="#FFFFFF">價格</font></td>
    <td width="141" align="center" ><font color="#FFFFFF">個數</font></td>
    <td width="178" align="center" ><font color="#FFFFFF">購買</font></td>
  </tr>
  <%
  rs.close
  rs.Open "SELECT 物品名,體力,內力,數量,銀兩,sm FROM 商店物品  where 擁有者="&shopname&" and 數量>0 and 類型='"&leixin&"'",conn
  do while not rs.bof and not rs.eof %>
  <tr> 
    <form method="POST" action='mai1.asp'>
      <td width="60" align="center" NOWRAP><font color="#000000"><%=rs("物品名")%></font></td>
      <td width="177" align="center"><font color="#000000">
        <img src="../hcjs/jhjs/001/<%=rs("sm")%>.gif" border="0" alt="<%=rs("物品名")%> ">
        </font></td>
      <td width="266" align="center"><font color="#000000">體力+<%=rs("體力")%> 內力+<%=rs("內力")%></font></td>
      <td width="67" align="center"><font color="#000000"><%=rs("數量")%></font></td>
      <td width="98" align="center"><font color="#000000"><%=rs("銀兩")%></font></td>
      <td width="141" align="center"> 
        <div align="center"> <font color="#000000"> 
          <input type="hidden" name="name" value=<%=rs("物品名")%>>
          <input type="hidden" name="shopname" value=<%=shopname%>>
          <input type="hidden" name="shang" value=<%=shang%>>
          <input type="hidden" name="money" value=<%=rs("銀兩")%>>
          <select  size="1" name="num">
            <option value="1" selected>1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="20">20</option>
            <option value="50">50</option>
            <option value="99">99</option>
          </select>
          </font></div>
      </td>
      <td width="178"> 
        <div align="center"> <font color="#000000"> 
          <input type="submit" value="購買" name="submit222222222222">
          </font></div>
      </td>
    </form>
  </tr>
  <% 
rs.MoveNext 
loop 
set rs=nothing
conn.close
set conn=nothing
%>
</table>
