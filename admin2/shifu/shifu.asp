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
sql="select 師傅,武功加,知質,內力,體力,性別,姓名,grade from 用戶 where ID=" & info(9)
set rs=conn.execute(sql) 
shifu=rs("師傅")
wushu=rs("武功加")
zhizi=rs("知質")
tili=rs("體力")
neili=rs("內力")%>
<html>

<head>
<title>隨師修煉</title>
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

<body BACKGROUND="../../ljimage/downbg.jpg" LEFTMARGIN="0" TOPMARGIN="0">
<table border="1" bgcolor="#007cd0" align="center" width="389" cellpadding="8"
cellspacing="8"> <tr bgcolor="#FFFFFF"> <td height="210"> <table border="1"
      width="355" cellspacing="2" cellpadding="1" bordercolor="#65251C"> <tr> 
<td colspan="6" height="33"> <div align="center"> <font size="2" class="c"><font size="3"><b><FONT COLOR="#0099FF" SIZE="2"><IMG SRC='../../ico/<%=info(10)%>-2.gif' WIDTH='54' 

HEIGHT='54'></FONT><font color="#FF0033"><%=rs("姓名")%></font></b><font
              size="2" color="#FF0033">大俠的江湖背景</font></font></font> </div></td></tr> 
<tr> <td width="113" bgcolor="#FFFFFF"><font size="2" class="c">姓 名</font></td><td colspan="5" bgcolor="#FFFFFF"><%=rs("姓名")%> 
江湖名號： <font color="#0000FF" size="2"> <%if rs("grade")<2  then response.write("初來乍道")               
if rs("grade")>=2and rs("grade")<3  then response.write("江湖初行")               
if rs("grade")>=3 and rs("grade")<4  then response.write("小有成就")               
if rs("grade")>=4 and rs("grade")<5  then response.write("聲名顯赫")               
if rs("grade")>=5 and rs("grade")<6  then response.write("行闖江湖")               
if rs("grade")>=6 and rs("grade")<7  then response.write("一代大俠")               
if rs("grade")>=7 and rs("grade")<8  then response.write("風雲劍客")               
if rs("grade")>=8 and rs("grade")<9  then response.write("聞名天下")               
if rs("grade")>=9 and rs("grade")<10  then response.write("雲遊仙勝")               
if rs("grade")>=10 then response.write("劍仙")               
            %> </font></td></tr> <tr> <td width="113" height="17" bgcolor="#FFFFFF"><font size="2" class="c">性 
別</font></td><td height="17" colspan="5" bgcolor="#FFFFFF"><%=rs("性別")%> </td></tr> 
<tr> <td width="113" bgcolor="#FFFFFF" height="10"><font size="2">提升的上限</font><font size="2" class="c">:</font></td><td width="51" bgcolor="#FFFFFF" height="10"><%=rs("武功加")%> 
</td><td width="96" bgcolor="#FFFFFF" height="10"><font size="2">資質：<%=rs("知質")%></font></td><td width="67" bgcolor="#FFFFFF" height="10">師傅：<font size="2"><%=rs("師傅")%></font></td></tr> 
<tr> <td width="113" bgcolor="#FFFFFF" height="6">內力：</td><td width="51" bgcolor="#FFFFFF" height="6"><%=rs("內力")%></td><td width="96" bgcolor="#FFFFFF" height="6">體力：<%=rs("體力")%></td><td width="67" bgcolor="#FFFFFF" height="6"><a href="liangong.asp">踏入練功房</a></td></tr> 
</table>【靈劍提示】有師傅的大俠可以在這和你的師傅學習武功， 學習一次，需要消耗體力100，資質20，提升<font color="#FF0000">練武的上限</font><br> 
10點，別小看這練武的資質哦。突破你的<font color="#FF3333">上限</font>全靠它了 <br> </table>

</body>

</html>
<%
conn.close
set conn=nothing
set rs=nothing%>
