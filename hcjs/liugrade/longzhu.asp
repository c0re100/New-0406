<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")


%>
<html>

<head>
<title>改良武器</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#660000" background="../../ff.jpg" LEFTMARGIN="0" TOPMARGIN="0">
<table border="1" bgcolor="#000000" align="center" width="389" cellpadding="8"
cellspacing="8">
  <tr bgcolor="#FFFFFF"> 
    <td height="99" bgcolor="#FFFF00"> 
      <table border="1"
      width="409" cellspacing="2" cellpadding="1" bordercolor="#65251C">
        <tr> 
          <td colspan="5" height="33" bgcolor="#000000"> 
            <div align="center"> 
              <p><font size="+2" color="#FFFF00">改良武器</font></p>
              <p><font size="2" color="#FFFFCC">條件：龍珠<img src="../../hcjs/jhjs/001/pig.gif" width="64" height="64" border="0" alt="點擊煉寶">可用來改良武器<br>
                等級越高改良成功機率越小！</font><font size="+2" color="#FFFFCC"> </font></p>
            </div>
          </td>
        </tr>
      </table>
      <table width="355" border="1" align="center" cellspacing="1" cellpadding="1" bordercolor="#FF0000">
        <tr> 
          <td width="121"> 
            <div align="center"> 
              <%rs.open "SELECT id,物品名,類型,等級,數量,攻擊,防御 FROM 物品 WHERE 擁有者=" & info(9) &" and 數量>0 and 裝備=false and (類型='頭盔' or 類型='左手' or 類型='右手' or 類型='雙腳' or 類型='裝飾' or 類型='盔甲')  order by 銀兩 ",conn%>
              <table width="395" border="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
                <tr bgcolor="#FF0000"> 
                  <td width="16%"> <div align="center"><font color="#FFFFFF">物品名</font></div></td>
                  <td width="8%"> <div align="center"><font color="#FFFFFF">類型</font></div></td>
                  <td width="12%"> <div align="center"><font color="#FFFFFF">等級</font></div></td>
                  <td width="10%"> <div align="center"><font color="#FFFFFF">數量</font></div></td>
                  <td width="10%"> <div align="center"><font color="#FFFFFF">加攻擊</font></div></td>
                  <td width="12%"> <div align="center"><font color="#FFFFFF">加防御</font></div></td>
                  <td width="12%"> <div align="center"><font color="#FFFFFF">操作</font></div></td>
                </tr>
                <%
do while not rs.eof
%>
                <tr align="center"> 
                  <td width="16%" height="15"><%=rs("物品名")%></td>
                  <td width="8%" height="15"> <%=rs("類型")%> </td>
                  <td width="12%" height="15"><%=rs("等級")%></td>
                  <td width="10%" height="15"> <%=rs("數量")%> </td>
                  <td width="10%" height="15"><%=rs("攻擊")%></td>
                  <td width="12%" height="15"><%=rs("防御")%></td>
                  <td width="12%" height="15"><a href="longzhuok.asp?id=<%=rs("ID")%>">改良</a></td>
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
            </div>
          </td>
          <td width="121"> 
            <div align="center"></div>
          </td>
        </tr>
      </table>
      <br>
      龍珠可在江湖裡打怪物隨機獲得，機率為80%<br>
      靈劍總站自創！ <br>
  </table>

</body>

</html>
