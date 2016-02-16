<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from 人命 where 時間<now()-3",conn,3,3
seekname=Trim(Request.Form("name"))
if seekname<>"" then
	rs.open "Select * From 人命 where 死者='"&seekname&"' or 兇手='"&seekname&"'Order by 時間 DESC",conn,3,3
else
	rs.open "Select * From 人命 Order by 時間 DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>江湖命案</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body bgcolor="#660000" background="../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<br>
<div align="center"> 
  <table width="550" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFCC00" bgcolor="#FFCC00">
    <tr> 
      <td width="100%"> 
        <table border="0" cellspacing="2" cellpadding="2" width="550" bordercolorlight="#EFEFEF" align="center">
          <tr bgcolor="#000066"> 
            <td width="20%"> 
              <div align="center"><font color="#FFFF99"> 死 亡 者 </font></div>
            </td>
            <td width="25%"> 
              <div align="center"><font color="#FFFF99"> 時 間 </font></div>
            </td>
            <td width="19%"> 
              <div align="center"><font color="#FFFF99"> 殺 人 兇 手 </font></div>
            </td>
            <td width="36%"> 
              <div align="center"><font color="#FFFF99"> 死 亡 原 因 </font></div>
            </td>
          </tr>
          <%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
          <tr bgcolor="#FF0000"> 
            <td width="20%"> 
              <div align="center"><%=rs("死者")%> </div>
            </td>
            <td width="25%"> 
              <div align="center"><%=rs("時間")%></div>
            </td>
            <td width="19%"> 
              <div align="center"><%=rs("兇手")%></div>
            </td>
            <td width="36%"> 
              <div align="center"><%=rs("死因")%></div>
            </td>
          </tr>
          <%rs.movenext%>
          <%count=count+1%>
          <%loop
			pa=rs.pagecount
			mu=musers()
          rs.close
          set rs=nothing
          conn.close
          set conn=nothing%>
        </table>
        <table border="0" cellspacing="1" cellpadding="1" width="550" bordercolorlight="#EFEFEF" align="center">
          <tr> 
            <td align="left" width="37%">[共<font color="red"><b><%=pa%></b></font>頁][<font
color="red"><b><%=mu%></b></font>人死亡]</td>
            <td align="right" width="63%"> 
              <div align="center">[<a href="listminan.asp?page=<%=page-1%>">上一頁</a>][<font color="#FFFFFF">第<%=page%>頁</font>][<a                    href="listminan.asp?page=<%=page+1%>">下一頁</a>]</div>
            </td>
          </tr>
        </table>
  </table>
</div>
<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 殺人 from 人命")
musers=tmprs("殺人")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>
</body> 
</html>