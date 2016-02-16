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
rs.open "Select lianbao from 集合 where id>=0 order by id",conn
lian=rs("lianbao")
rs.close
sql="select 會員等級,寶物,法力,知質 from 用戶 where ID=" & info(9)
set rs=conn.execute(sql) 
zhizi=rs("知質")
if rs("會員等級")=0 and lian=1 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你不是會員請回吧！');location.href = '../help/info.asp?type=2&name=會員加入辦法';}</script>"
Response.End
end if
if zhizi<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的資質低於1000無此能力煉寶！');window.close();}</script>"
Response.End
end if
if rs("寶物")="無" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你哪來的寶物靈劍水晶石呀！');window.close();}</script>"
Response.End
end if

%>
<html>

<head>
<title>煉寶</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css"><!--td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { text-decoration: none; font-family: 新細明體; font-size: 9pt }
a:hover      { text-decoration: underline; color: #CC0000; font-family: 新細明體; font-size: 
               9pt }
.big         { font-family: 新細明體; font-size: 12pt }
.txt         { font-family: 新細明體; font-size: 10.8pt }
--></style>
</head>

<body bgcolor="#660000" LEFTMARGIN="0" TOPMARGIN="0">
<table border="1" bgcolor="#003300" align="center" width="389" cellpadding="8"
cellspacing="8">
  <tr bgcolor="#FFFFFF"> 
    <td height="99" bgcolor="#99CCFF"> 
      <table border="1"
      width="355" cellspacing="2" cellpadding="1" bordercolor="#65251C">
        <tr> 
          <td colspan="5" height="33"> 
            <div align="center"> <b><font size="+2" color="#FF0000">煉寶</font></b></div>
          </td>
        </tr>
        <tr> 
          <td width="113" height="17" bgcolor="#FFFFFF"><font size="2" class="c">寶物：</font></td>
          <td height="17" colspan="4" bgcolor="#FFFFFF"><%=rs("寶物")%> </td>
        </tr>
        <tr> 
          <td width="113" bgcolor="#FFFFFF" height="10"><font size="2">當前法力</font><font size="2" class="c">:</font></td>
          <td width="51" bgcolor="#FFFFFF" height="10"><%=rs("法力")%> </td>
          <td bgcolor="#FFFFFF" height="10"><font size="2">資質：<%=rs("知質")%></font></td>
        </tr>
      </table>
      <table width="355" border="1" align="center" cellspacing="1" cellpadding="1" bgcolor="#000000" bordercolor="#CCCCCC">
        <tr> 
          <td width="121"> 
            <div align="center"><a href="liangong.asp"><font color="#FFFFFF">聚<br>
              精<br>
              會<br>
              神</font></a></div>
          </td>
          <td width="89"><a href="lianbaook.asp"><img src="111.gif" width="79" height="75" border="0" alt="點擊煉寶"></a></td>
          <td width="121"> 
            <div align="center"><font color="#FFFFFF">心<br>
              無<br>
              旁<br>
              念</font></div>
          </td>
        </tr>
      </table>
      <br>
      目前隻對會員開放，你有了寶物（靈劍水晶石）後可以在此煉寶吸取寶物的法力,獲得的法力多少根據你的資質而定！<br>
    
  </table>

</body>

</html>
<%
conn.close
set conn=nothing
set rs=nothing%>
