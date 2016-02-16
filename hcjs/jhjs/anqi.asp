<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>道具店</title>
<link rel="stylesheet" href="../../css.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#003399" background="../../ff.jpg"
leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> <td align="center" valign="top"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2">
        <tr bgcolor="#FFFF00"> 
          <td width="17%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a href="BINQI.asp">手持刀劍</a> 
              </font></div></td>
          <td width="16%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a href="dunpai.asp">手持護具</a></font></div></td>
          <td width="16%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a href="body.asp">盔甲</a></font></div></td>
          <td width="12%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a href="head.asp">頭盔</a></font></div></td>
          <td width="13%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a href="foot.asp">雙腳</a> 
              </font></div></td>
          <td width="11%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a href="other.asp">裝飾 
              </a></font></div></td>
          <td width="15%" nowrap> 
            <div align="center"><font color="#FFFFFF"><a
href="anqi.asp">暗器</a></font></div></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2" bgcolor="#99CCFF">
        <tr align="center" bgcolor="#CC0000"> 
          <td width="69"><font color="#FFFFFF">裝備名</font></td>
          <td width="85"><font color="#FFFFFF"> 武 器 圖</font></td>
          <td width="58"><font color="#FFFFFF">類型 </font> 
          <td width="244"><font color="#FFFFFF">功能</font></td>
          <td width="132"><font color="#FFFFFF">售價</font></td>
          <td colspan="2"><font color="#FFFFFF">數量</font></td>
          <td width="46"><font color="#FFFFFF">操作</font></td>
        </tr>
        <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 物品名,說明,id,類型,內力,體力,堅固度,銀兩 FROM 物品買 where 類型='暗器' order by 銀兩 ",conn
do while not rs.bof and not rs.eof
%>
        <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
          <form method=POST action='binqi3.asp?id=<%=rs("id")%>'>
            <td width="69"> <div align="center"><%=rs("物品名")%></div></td>
            <td width="85"> <div align="center"><img src="001/<%=rs("說明")%>.gif" border="0" alt="<%=rs("物品名")%> "></div></td>
            <td width="58"> <div align="center"><%=rs("類型")%></div></td>
            <td width="244"> <div align="center">傷內力<%=abs(rs("內力"))%>，傷體力<%=abs(rs("體力"))%>，堅固度<%=abs(rs("堅固度"))%></div></td>
            <td width="132"> <div align="right"><%=rs("銀兩")%>兩</div></td>
            <td colspan="2"> <div align="center"> 
                <select name="sl" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
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
              </div></td>
            <td width="46"> <div align="center"> 
                <input type="SUBMIT" name="Submit"
value="進行">
              </div></td>
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
    </td></tr> </table>
</body>
</html>
