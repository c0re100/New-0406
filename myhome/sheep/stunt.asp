<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<title>寵物技能商店</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" background="../../ff.jpg" leftmargin="0" topmargin="7">
<div align="left"></div>
<div align="center">
  <table width="480" border="0" cellspacing="0" cellpadding="4">
    <tr> 
      <td colspan="2" valign="top" align="center"> [ <font color="#009900">寵 物 
        商 店</font> ]<br> <br>
        【<a href="indexsheep.asp"> 寵物商店 </a>】 【<a href="stunt.asp"> 技能商店</a> 】 
        【 <a href="itemshop.asp">道具商店 </a>】【<a href="sellsheep.asp"> <font color="#FF6600">我要賣掉我的寵物</font></a> 
        】 <br> <br> <div align="center"> </div>
        <div align="center"><font size="-1">歡迎光臨寵物商店，這裡有各式不同種類的寵供你選購呵。注意，寵物隻能買一隻呵。</font><br>
          <!--#include file="data.asp"-->
          <br>
          <table border="1" cellspacing="2" cellpadding="0" align="center" bordercolor="#FF0000" width="600">
            <tr bgcolor="#FF0000"> 
              <td width="109"> <div align="center"><font color="#FFFFFF">技能名稱 </font></div></td>
              <td width="190"> <div align="center"><font color="#FFFFFF">效 果 </font></div></td>
              <td width="114"> <div align="center"><font color="#FFFFFF">售 價 </font></div></td>
              <td width="167"> <div align="center"><font color="#FFFFFF">操 作 </font></div></td>
            </tr>
            <%
sql="SELECT 技能名稱,簡介,價格 FROM 寵物技能'"
Set rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
            <tr> 
              <td><%=rs("技能名稱")%></td>
              <td><%=rs("簡介")%> </td>
              <td> <div align="center"><%=rs("價格")%>兩</div></td>
              <td> <p align="center"><a href="buystunt.asp?name=<%=rs("技能名稱")%>&money=<%=rs("價格")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></p></td>
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
          <br>
        </div></td>
    </tr>
  </table>
</div>
</body>

</html>