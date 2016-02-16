<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="Select  武功加,內力加,攻擊加,防御加,體力加,魅力加,道德加,法力點,姓名,狀態,性別,會員費,門派,體力,等級,身份,內力,存款,武功,銀兩,攻擊,配偶,防御,二婚,道德,介紹人,魅力,grade from 用戶 where 姓名='"&info(0)&"'"
set rs=conn.Execute(sql)
wgj=rs("武功加")
nlj=rs("內力加")
gjj=rs("攻擊加")
fyj=rs("防御加")
tlj=rs("體力加")
mlj=rs("魅力加")
ddj=rs("道德加")
flj=rs("法力點")
%>
<html>

<head>
<title>江湖檔案</title>
<LINK href="../css.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</head>

<body topmargin="0" leftmargin="0" background="../images/8.jpg" text="#993333" vlink="#663366" link="#660066">
<table width="412" border="0" cellspacing="0" cellpadding="0" height="310" align="center">
  <tr>
    <td width="412" height="308"> 
      <table border="1"
width="407" cellspacing="1" cellpadding="1" align="center">
        <tr> 
          <td colspan="5" height="23"> 
            <table border="0"
width="400" cellspacing="1" cellpadding="1" align="center" background="../ljimage/downbg.jpg">
              <tr> 
                <td width="65" height="25"><font color="#0099FF" size="2"><img src='../ico/<%=info(10)%>-2.gif' width='54' 

height='54'></font></td>
                <td width="111" height="25"><font color="#0099FF" size="2"><%=rs("姓名")%></font></td>
                <td width="63" height="25">狀 態：</td>
                <td height="25" colspan="2" width="148"> <font color="#0099FF" size="2"><%=rs("狀態")%></font></td>
              </tr>
              <tr> 
                <td width="65" height="25">性 別：</td>
                <td width="111" height="25"><font color="#0099FF" size="2"><%=rs("性別")%></font></td>
                <td width="63" height="25">會 費：</td>
                <td height="25" colspan="2" width="148"> <font color="#0099FF" size="2"><%=rs("會員費")%></font><font size="2">元</font></td>
              </tr>
              <tr> 
                <td width="65" height="25">門 派：</td>
                <td width="111" height="25"><%=rs("門派")%></td>
                <td width="63" height="25">體 力：</td>
                <td height="25" colspan="2" width="148"><%=rs("體力")%> / <font color="#0099FF" size="2"><%=rs("等級")%></font></td>
              </tr>
              <tr> 
                <td width="65" height="20">身 份：</td>
                <td width="111" height="24"><%=rs("身份")%></td>
                <td width="63" height="24">內 力：</td>
                <td height="24" colspan="2" width="148"><%=rs("內力")%> / <font color="#0099FF" size="2"><%=rs("等級")*640+2000+nlj%></font> 
                </td>
              </tr>
              <tr> 
                <td width="65" height="20">存 款：</td>
                <td width="111" height="24"><%=rs("存款")%> 兩</td>
                <td width="63" height="24">武 功：</td>
                <td height="24" colspan="2" width="148"><%=rs("武功")%> / <font color="#0099FF" size="2"><%=rs("等級")*1280+3800+wgj%> 
                  </font> </td>
              </tr>
              <tr> 
                <td width="65" height="20">現 金：</td>
                <td width="111" height="24"><%=rs("銀兩")%> 兩</td>
                <td width="63" height="24">攻 擊：</td>
                <td height="24" colspan="2" width="148"><%=rs("攻擊")%> / <font color="#0099FF" size="2"><%=rs("等級")*1200+4500+gjj%> 
                  </font> </td>
              </tr>
              <tr> 
                <td width="65" height="20">配 偶：</td>
                <td width="111"><%=rs("配偶")%> </td>
                <td width="63">防 御：</td>
                <td colspan="2" width="148"><%=rs("防御")%> / <font color="#0099FF" size="2"><%=rs("等級")*1120+3000+fyj%> 
                  </font> </td>
              </tr>
              <tr> 
                <td width="65" height="20">二 婚：</td>
                <td width="111"><%=rs("二婚")%></td>
                <td width="63">道 德：</td>
                <td colspan="2" width="148"><%=rs("道德")%><font size="2"> /</font> 
                  <font color="#0099FF" size="2"><%=rs("等級")*1440+2200+ddj%></font> 
                </td>
              </tr>
              <tr> 
                <td width="65" height="20">介紹人：</td>
                <td width="111"><%=rs("介紹人")%> </td>
                <td width="63">魅 力：</td>
                <td colspan="2" width="148"><%=rs("魅力")%> <%=rs("等級")*1120+2100+mlj%> 
                </td>
              </tr>
              <tr> 
                <td width="65" height="20">賭場等級：</td>
                <td width="111">賭神</td>
                <td width="63">戰鬥級：</td>
                <td colspan="2" width="148"><%=rs("等級")%> 級</td>
              </tr>
              <tr> 
                <td width="65" height="20">贏 / 賭：</td>
                <td width="111"> </td>
                <td width="63">管理級：</td>
                <td colspan="2" width="148"><%=rs("grade")%> 級</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
      
      </td>
</tr>
</table>
</body>

</html>
<%
rs.close
set rs=nothing
%> 