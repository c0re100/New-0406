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
<title>藥品店</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#003399" background="../../ff.jpg"
leftmargin="0" topmargin="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> <td align="center" valign="top"> 
      <table border="0" cellpadding="2" cellspacing="2" width="100%" bgcolor="#99CCFF">
        <tr bgcolor="#FFFF00"> 
          <td height="13" colspan="6"> 
            <div align="center"></div>
            <div align="center">[ <font color="#FF3300">補 藥 專 賣 店</font> ]--[ 
              <a
href="duyao.asp"><font color="#6633FF">毒 藥 專 賣 店</font></a> ]</div></td>
        </tr>
        <tr> 
          <td width="124" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">補 
              藥</font></div></td>
          <td width="99" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF"> 
              藥 品 圖</font></div></td>
          <td width="266" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">功 
              能</font></div></td>
          <td width="91" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">售 
              價</font></div></td>
          <td width="85" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">數 
              量</font></div></td>
          <td width="92" height="13" bgcolor="#FF0000"> 
            <div align="center"><font color="#FFFFFF">操 
              作</font></div></td>
        </tr>
        <%
rs.open "SELECT id,物品名,說明,內力,體力,銀兩 FROM 物品買 where  類型='藥品' order by 銀兩",conn
do while not rs.bof and not rs.eof
%>
        <tr  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
          <form method=POST action='buyyao.asp?id=<%=rs("id")%>'>
            <td width="124" bgcolor="#FFFFFF"> 
              <div align="center"><%=rs("物品名")%></div></td>
            <td width="99" bgcolor="#FFFFFF"> 
              <div align="center"><img src="001/<%=rs("說明")%>.gif" border="0" alt="<%=rs("物品名")%> "></div></td>
            <td width="266" bgcolor="#FFFFFF">補內力<%=rs("內力")%>，補生命<%=rs("體力")%>!</td>
            <td width="91" bgcolor="#FFFFFF"> 
              <div align="center"><%=rs("銀兩")%>兩</div></td>
            <td width="85" bgcolor="#FFFFFF"> 
              <div align="right"> 
                <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="99">99</option>
                </select>
              </div></td>
            <td width="92" bgcolor="#FFFFFF"> 
              <div align="center"> 
                <input type="SUBMIT" name="Submit" value="進行">
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
    </td></tr> </table><br> 
</body>

</html>
