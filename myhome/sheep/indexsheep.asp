<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<%info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
my=info(0)
%>
<html>
<head>
<title>寵物商店區</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" leftmargin="0" topmargin="7">
<div align="left"></div>
<div align="center">
  <table width="640" border="0" cellspacing="0" cellpadding="4">
    <tr> 
      <td width="615" colspan="2" align="center" valign="top"> [ <font color="#009900">寵 物 
        商 店</font> ]<br> <br>
        【<a href="indexsheep.asp"> 寵物商店 </a>】 【<a href="stunt.asp"> 技能商店</a> 】 
        【 <a href="itemshop.asp">道具商店 </a>】 【<a href="sellsheep.asp"> <font color="#FF6600">我要賣掉我的寵物</font></a> 
        】 <br> <br> <div align="center"> </div>
        <div align="center"><font size="-1">歡迎光臨寵物商店，這裡有各式不同種類的寵供你選購呵。注意，寵物隻能買一隻呵。</font><br>
          <!--#include file="data.asp"-->
          <table width="620" border="1" cellspacing="2" cellpadding="0" bordercolor="#FF0000" height="26">
            <tr bgcolor="#FF0000"> 
              <td width="145" height="17"> <div align="center"><font color="#FFFFFF">寵物類型 
                  </font></div></td>
              <td width="72" height="17"> <div align="center"><font color="#FFFFFF">寵物技能</font></div></td>
              <td width="187" height="17"> <div align="center"><font color="#FFFFFF">寵物參數 
                  </font></div></td>
              <td width="72" height="17"> <div align="center"><font color="#FFFFFF">售 
                  價 </font></div></td>
              <td width="72" height="17"> <div align="center"><font color="#FFFFFF">操 
                  作 </font></div></td>
            </tr>
            <%
sql="SELECT 寵物類型,技能,說明,攻擊,防御,體力,價格 FROM 寵物'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
            <tr> 
              <td width="145" height="22"><img src="image/<%=rs("說明")%>.gif" border="0" alt="<%=rs("寵物類型")%> " align="absmiddle"></td>
              <td width="72" height="22"> <div align="center"><%=rs("技能")%></div></td>
              <td width="187" height="22">攻擊：<%=rs("攻擊")%> 防御：<%=rs("防御")%> 體力：<%=rs("體力")%> </td>
              <td width="72" height="22"> <div align="center"><%=rs("價格")%>兩</div></td>
              <td width="72" height="22"> <p align="center"><a href="buy.asp?name=<%=rs("寵物類型")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></p></td>
            </tr>
            <%
rs.movenext
loop
conn.close
%>
          </table>
          <br>
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>