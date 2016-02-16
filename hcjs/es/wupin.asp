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
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,物品名,類型,數量,內力,體力,攻擊,防御,銀兩 FROM 物品 WHERE 擁有者="& info(9) & " and 數量>0 and 裝備=false and 類型<>'卡片' order by 類型 ",conn
%>
<html>
<head>
<title>物品管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#660000" background="../../chat/bk.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<div align="left">
  <div align="center"> </div>
  <div align="center">
    <table border="0" align="center" width="640" cellpadding="2" cellspacing="2" height="31">
      <tr align="center" bgcolor="#990000"> 
        <td width="53" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF" size="-1">物品名</font></div></td>
        <td width="30" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">類型 
            </font></div>
        <td width="40" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF" size="-1">數量 
            </font> </div>
        <td width="41" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">內力 
            </font></div>
        <td width="42" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">體力 
            </font></div>
        <td width="39" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">攻擊 
            </font></div>
        <td width="42" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">防御 
            </font></div>
        <td height="16" colspan="2" nowrap> 
          <div align="center"><font color="#FFFFFF" size="-1">價錢</font></div></td>
        <td width="56" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">方式</font></div></td>
        <td width="48" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">數量</font></div></td>
        <td width="54" height="16" nowrap> 
          <div align="center"><font color="#FFFFFF">操作</font></div></td>
      </tr>
      <%
do while not rs.eof
%>
      <tr bgcolor="#FFFFFF"> 
        <form method=POST action='fs.asp?id=<%=rs("id")%>&name=<%=name%>'>
          <td width="53" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("物品名")%> </font> </div></td>
          <td width="30" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("類型")%></font> </div>
          <td width="40" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("數量")%> </font> </div>
          <td width="41" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("內力")%></font> </div>
          <td width="42" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("體力")%></font> </div>
          <td width="39" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("攻擊")%></font> </div>
          <td width="42" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("防御")%></font> </div>
          <td colspan="2" height="3"> <div align="center"><font color="#000000" size="-1"><%=rs("銀兩")%> </font> </div></td>
          <td height="3" width="56"> <div align="center"><font color="#FFFFFF"> 
              <select name="wpfs">
                <option value="1" selected>&nbsp;贈送</option>
                <option value="2">&nbsp;交易</option>
                <option value="3">&nbsp;二手</option>
                <option value="4">保險櫃</option>
              </select>
              </font></div></td>
          <td height="3" width="48"> <div align="center"> <font color="#FFFFFF"> 
              <select name="wpsl">
                <option value="1" selected>1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
              </select>
              </font></div></td>
          <td height="3" width="54"> <div align="center"><font color="#FFFFFF"> 
              <input type="SUBMIT" name="Submit"  value="進行">
              </font></div></td>
        </form>
      </tr>
      <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
    </table>    
    <table width="640" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><font color="#FFFFFF">說明：在這時你可以對自己的物品進行管理操作。<br>
          1.贈送：把自己的東西贈送給她（他），此操作隻是發生在異性之間。可以用此進行交流。<br>
          2.交易：你可以與你的朋友進行物品交易，別人是買不到的。價錢由你定，親兄弟明算賬嘛。<br>
          3.二手：這裡你可以漫天要價，一個願賣，一個願買。不過讓人“砸”了大頭，我們可不管。<br>
          4.保險櫃：有什麼東西想自己偷偷存起來，這裡可是好地方呀。是不會丟的，不過收費喲。</font></td>
</tr>
</table>
</div>
</div>
</body>
</html>
