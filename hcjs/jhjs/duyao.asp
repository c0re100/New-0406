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
<title>江湖藥鋪</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#003399" background="../../ff.jpg"
leftmargin="0" topmargin="0">
<div align="center"> 
  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td align="center" valign="top"> <table width="100%" border="0" align="center" cellpadding="2" cellspacing="2">
          <tr> 
            <td colspan="6" bgcolor="#FFFF00"> <div align="center"> 
                <div align="center"></div>
                <div align="center">[ <font color="#FF3300"><a href="yaopu0.asp"><font color="#6633FF">補 
                  藥 專 賣 店</font></a></font> ]--[ <font color="#FF0000">毒 藥 專 賣 
                  店</font> ]</div>
              </div></td>
          </tr>
          <tr bgcolor="#FF0000"> 
            <td width="125"> <div align="center"><font color="#FFFFFF">毒 藥</font></div></td>
            <td width="115" height="13"> <div align="center"><font color="#FFFFFF"> 
                藥 品 圖</font></div></td>
            <td width="244"> <div align="center"><font color="#FFFFFF">功 能</font></div></td>
            <td width="76"> <div align="center"><font color="#FFFFFF">售 價</font></div></td>
            <td width="93"> <div align="center"><font color="#FFFFFF">數 量</font></div></td>
            <td width="104"> <div align="center"><font color="#FFFFFF">操 作</font></div></td>
          </tr>
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id,物品名,內力,體力,銀兩,說明 FROM 物品買 where  類型='毒藥' order by 銀兩",conn
do while not rs.bof and not rs.eof
%>
          <tr  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
            <form method=POST action='buyduyao.asp?id=<%=rs("id")%>'>
              <td width="125" bgcolor="#FFFFFF"> <div align="center"><%=rs("物品名")%></div></td>
              <td width="115" bgcolor="#FFFFFF"> <div align="center"><img src="001/<%=rs("說明")%>.gif" border="0" alt="<%=rs("物品名")%> "></div></td>
              <td width="244" bgcolor="#FFFFFF">失內力<%=rs("內力")%>，失生命<%=rs("體力")%></td>
              <td width="76" bgcolor="#FFFFFF"> <div align="center"><%=rs("銀兩")%>兩</div></td>
              <td width="93" bgcolor="#FFFFFF"> <div align="center"> 
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
              <td width="104" bgcolor="#FFFFFF"> <div align="center"> 
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
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
