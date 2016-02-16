<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
tiaojian=Request.Form("tiaojian")
show=trim(Request.Form("show"))
fs=int(Request.form("seekfs"))
%>
<html>
<head>
<title>用戶資料查看程序</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="0064.jpg">
<p align="center"> <font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font> 
  <br>
  這一些是滿足條件的人！點姓名進行修改！<br>
  <font color="#FF0000"><b><%=tiaojian%></b></font> <br>
 <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
select case fs
case 0
	if show<>"" then
		rs.open "SELECT * FROM 用戶 where "& tiaojian &" order by -"&show,conn
	else
		rs.open "SELECT * FROM 用戶 where "& tiaojian &" order by lasttime",conn
	end if
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center">
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">靈劍江湖</font></b></font></td>
</tr>
<tr align="center">
<td>
      <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr> 
          <td width="28" height="10"> 
            <div align="center"><font color="#FFFFFF">ID</font></div>
          </td>
          <td width="47" height="10"> 
            <div align="center"><font color="#FFFFFF">姓名</font></div>
          </td>
          <td width="25" height="10"> 
            <div align="center"><font color="#FFFFFF">性別</font></div>
          </td>
          <td width="63" height="10"> 
            <div align="center"><font color="#FFFFFF">門派</font></div>
          </td>
          <td width="54" height="10"> 
            <div align="center"><font color="#FFFFFF">身份</font></div>
          </td>
          <td width="82" height="10"> 
            <div align="center"><font color="#FFFFFF">最後登陸時間</font></div>
          </td>
          <td width="58" height="10"> 
            <div align="center"><font color="#FFFFFF">江湖等級</font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center"><font color="#FFFF00"><b><%=show%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center"><font color="#FFFFFF">登錄</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="28" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
          </td>
          <td width="47" height="30"> 
            <div align="center"><a href="SHOWUSER.asp?username=<%=rs("姓名")%>"><font color="#FF9900"><%=rs("姓名")%></font></a></div>
          </td>
          <td width="25" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("性別")%></font></div>
          </td>
          <td width="63" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("門派")%></font></div>
          </td>
          <td width="54" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("身份")%></font></div>
          </td>
          <td width="82" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("lasttime")%></font></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("等級")%></font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><font color="#FFFF00"><b><%=rs(show)%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
conn.close
%>
      </table>
</td>
</tr>
</table>
<div align="center"><font color="#000000">條件人數:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font> 
</div>
<% case 1
sql="SELECT * FROM 物品 where "& tiaojian &" order by 類型,數量"
Set Rs=conn.Execute(sql)
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center"> 
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">靈劍江湖</font></b></font></td>
  </tr>
  <tr align="center"> 
    <td> 
      <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr> 
          <td width="26" height="7"> 
            <div align="center"><font color="#FFFFFF">ID</font></div>
          </td>
          <td width="41" height="7"> 
            <div align="center"><font color="#FFFFFF">物品名</font></div>
          </td>
          <td width="52" height="7"> 
            <div align="center"><font color="#FFFFFF">擁有者</font></div>
          </td>
          <td width="32" height="7"> 
            <div align="center"><font color="#FFFFFF">類型</font></div>
          </td>
          <td width="110" height="7"> 
            <div align="center"><font color="#FFFFFF">說明</font></div>
          </td>
          <td width="34" height="7"> 
            <div align="center"><font color="#FFFFFF">裝備</font></div>
          </td>
         
          <td width="53" height="7"> 
            <div align="center"><b><font color="#FFFFFF">銀兩</font></b></div>
          </td>
          <td width="58" height="7"> 
            <div align="center"><font color="#FFFFFF">數量</font></div>
          </td>
         
          <td width="59" height="7"> 
            <div align="center"><font color="#FFFFFF">操作</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="26" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
          </td>
          <td width="41" height="30"> 
            <div align="center"><font color="#FF9900"><%=rs("物品名")%></font></div>
          </td>
          <td width="52" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("擁有者")%></font></div>
          </td>
          <td width="32" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("類型")%></font></div>
          </td>
          <td width="110" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("說明")%></font></div>
          </td>
          <td width="34" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("裝備")%></font></div>
          </td>
          
          <td width="53" height="30"> 
            <div align="center"><b><font color="#FFFFFF"><%=rs("銀兩")%></font></b></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("數量")%></font></div>
          </td>
         
          <td width="59" height="30"> 
            <div align="center"><font size="-1"><a href="../chat/del.asp?id=<%=rs("id")%>"><font color="#FFFF00">刪除</font></a></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
'conn.close
 case 2
sql="SELECT * FROM 修練物品 where "& tiaojian &" order by 數量"
Set Rs=conn.Execute(sql)
%>
      </table>
    </td>
  </tr>
</table>
<div align="center"><font color="#000000">條件人數:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font> 
  <br>
  <table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
    <tr align="center"> 
      <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">靈劍江湖</font></b></font></td>
    </tr>
    <tr align="center"> 
      <td> 
        <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
          <tr> 
            <td width="32" height="12"> 
              <div align="center"><font color="#FFFFFF">ID</font></div>
            </td>
            <td width="64" height="12"> 
              <div align="center"><font color="#FFFFFF">物品名</font></div>
            </td>
            <td width="71" height="12"> 
              <div align="center"><font color="#FFFFFF">擁有者</font></div>
            </td>
            <td width="61" height="12"> 
              <div align="center"><font color="#FFFFFF">數量</font></div>
            </td>
            <td width="133" height="12"> 
              <div align="center"><font color="#FFFFFF">功效</font></div>
            </td>
            <td width="55" height="12"> 
              <div align="center"><font color="#FFFFFF">增加數值</font></div>
            </td>
            <%if show<>"" then%>
            <%end if%>
            <td width="53" height="12"> 
              <div align="center"><font color="#FFFFFF">操作</font></div>
            </td>
          </tr>
          <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
          <tr> 
            <td width="32" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
            </td>
            <td width="64" height="16"> 
              <div align="center"><font color="#FF9900"><%=rs("物品名")%></font></div>
            </td>
            <td width="71" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("擁有者")%></font></div>
            </td>
            <td width="61" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("數量")%></font></div>
            </td>
            <td width="133" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("功效")%></font></div>
            </td>
            <td width="55" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("增加數值")%></font></div>
            </td>
            <%if show<>"" then%>
            <%end if%>
            <td width="53" height="16"> 
              <div align="center"><font size="-1"><a href="../chat/del1.asp?id=<%=rs("id")%>"><font color="#FFFF00">刪除</font></a></font></div>
            </td>
          </tr>
          <%
rs.movenext
loop
conn.close
%>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"><font color="#000000">條件人數:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">人</font> 
  </div>
  <br>
</div>
<%end select%>
</body>
</html> 