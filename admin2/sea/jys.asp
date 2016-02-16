<%
Response.Expires=0
Response.Buffer=True
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
action=Request("action")
if action="true" then
if instr(Request.ServerVariables("HTTP_REFERER"),Request.ServerVariables("SERVER_NAME"))<>0 then
port=Request.Form("port")
what=Request.Form("type")
submit=trim(Request.Form("submit"))
num=trim(Request.Form("num"))
On Error Resume Next
num=Clng(num)
if err.number>0 then
  err.clear
  response.write "<script language='javascript'>alert('請輸入數字！！！');history.back(-1);</script>"
  response.end
end if
if num<1 or num=empty then num=1
if num>10 then num=10
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn.Open Connstr
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open Application("ljjh_usermdb")
Set RS=Conn.Execute("select 擁有者,次數,時間 from 航海物品 where 擁有者='"& name &"'")
if RS.eof and RS.bof then
  Conn.Execute("insert into 航海物品(擁有者) values('"& name &"')")
  times=8
else
times=rs("次數")
  if now()-rs("時間")>1 then
    conn.execute("update 航海物品 set 次數=次數+800,時間=now() where 擁有者='"& name &"'")
    times=times+800
  end if
end if
RS.close
if times<1 then
mess="你的行動力沒了！！！"
else
Rs=Conn1.execute("select 銀兩 from 用戶 where 姓名='"& name &"'")
yin=Rs("銀兩")
Sql="select * from 航海時代 where 港口='"& port &"'"
Rs=Conn.Execute(Sql)
jia1=Rs("絲綢") : jia2=Rs("人參") : jia3=Rs("珠寶") : jia4=Rs("瓷器")
  if submit="買" then
    Select Case what
      Case "絲綢"
          if num*jia1>yin then
            mess="你的錢不夠了！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩-"& num*jia1 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,絲綢=絲綢+"& num &" where 擁有者='"& name &"'")
            Conn.Execute("update 航海時代 set 絲綢=絲綢+1 where 港口='"& port &"'")
            mess="成功購買！！！"
          end if
      Case "人參"
          if num*jia2>yin then
            mess="你的錢不夠了！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩-"& num*jia2 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,人參=人參+"& num &" where 擁有者='"& name &"'")
            Conn.Execute("update 航海時代 set 人參=人參+2 where 港口='"& port &"'")
            mess="成功購買！！！"
          end if
      Case "珠寶"
          if num*jia3>yin then
            mess="你的錢不夠了！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩-"& num*jia3 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,珠寶=珠寶+"& num &" where 擁有者='"& name &"'")
            Conn.Execute("update 航海時代 set 珠寶=珠寶+3 where 港口='"& port &"'")
            mess="成功購買！！！"
          end if
      Case "瓷器"
          if num*jia4>yin then
            mess="你的錢不夠了！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩-"& num*jia4 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,瓷器=瓷器+"& num &" where 擁有者='"& name &"'")
            Conn.Execute("update 航海時代 set 瓷器=瓷器+1 where 港口='"& port &"'")
            mess="成功購買！！！"
          end if
     End Select
  else
    Select Case what
      Case "絲綢"
          rs=Conn.Execute("select 絲綢 from 航海物品 where 擁有者='"& name &"'")
          if rs("絲綢")<num then
            mess="你沒那麼多絲綢！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩+"& num*jia1 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,絲綢=絲綢-"& num &" where 擁有者='"& name &"'")
            if jia1>1 then
               Conn.Execute("update 航海時代 set 絲綢=絲綢-1 where 港口='"& port &"'")
            end if
            mess="成功賣出！！！"
          end if
      Case "人參"
          rs=Conn.Execute("select 人參 from 航海物品 where 擁有者='"& name &"'")
          if rs("人參")<num then
            mess="你沒那麼多人參！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩+"& num*jia2 &" where 姓名='"& my &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,人參=人參-"& num &" where 擁有者='"& name &"'")
            if jia1>2 then
               Conn.Execute("update 航海時代 set 人參=人參-2 where 港口='"& port &"'")
            end if     
            mess="成功賣出！！！"
          end if
      Case "珠寶"
          rs=Conn.Execute("select 珠寶 from 航海物品 where 擁有者='"& name &"'")
          if rs("珠寶")<num then
            mess="你沒那麼多珠寶！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩+"& num*jia3 &" where 姓名='"& my &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,珠寶=珠寶-"& num &" where 擁有者='"& name &"'")
            if jia1>3 then
               Conn.Execute("update 航海時代 set 珠寶=珠寶-3 where 港口='"& port &"'")
            end if
            mess="成功賣出！！！"
          end if
      Case "瓷器"
          rs=Conn.Execute("select 瓷器 from 航海物品 where 擁有者='"& name &"'")
          if rs("瓷器")<num then
            mess="你沒那麼多瓷器！！！"
          else
            Conn1.Execute("update 用戶 set 銀兩=銀兩+"& num*jia4 &" where 姓名='"& name &"'")
            Conn.Execute("update 航海物品 set 次數=次數-1,瓷器=瓷器-"& num &" where 擁有者='"& name &"'")
            if jia1>1 then
               Conn.Execute("update 航海時代 set 瓷器=瓷器-1 where 港口='"& port &"'")
            end if
            mess="成功賣出！！！"
          end if
     End Select
  end if
end if
set Rs=nothing
Conn.close : set Conn=nothing
Conn1.close : set Conn1=nothing
response.write "<script language='javascript'>alert('"& mess &"');history.back(-1);</script>"
else
response.write "請不要作弊！！！"
end if
else
port=Request.Querystring("port")
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn.Open Connstr
Sql="select * from 航海時代 where 港口='"& port &"'"
Set Rs=Conn.Execute(Sql)
%>
<html>
<head>
<title><%=port%>港商會</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "新細明體"; font-size: 12px}
input {  font-family: "新細明體"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<table width="90%" border="1" cellspacing="2" cellpadding="2" align="center">
  <tr align="center" bgcolor="#FFFFFF"> 
    <td colspan="4" height="22"><%=port%>港商會 <br>
      <br>
      (一次至多購買10單位)<br>
    </td>
  </tr>
  <tr align="center" bgcolor="#660000"> 
    <td height="20"><font color="#FFFFFF">絲綢</font></td>
    <td height="20"><font color="#FFFFFF">人參</font></td>
    <td height="20"><font color="#FFFFFF">珠寶</font></td>
    <td height="20"><font color="#FFFFFF">瓷器</font></td>
    <%do while not Rs.eof%> </tr>
  <tr align="center" bgcolor="#FFFFFF"> 
    <td height="20"><%=Rs("絲綢")%>兩/單位</td>
    <td height="20"><%=Rs("人參")%>兩/單位</td>
    <td height="20"><%=Rs("珠寶")%>兩/單位</td>
    <td height="20"><%=Rs("瓷器")%>兩/單位</td>
  </tr>
  <tr align="center">
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 買 " OnClick=check()>
        <input type="submit" name="submit" value=" 賣 " OnClick=check()>
        <input type="hidden" name="type" value="絲綢">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 買 " OnClick=check()>
        <input type="submit" name="submit" value=" 賣 " OnClick=check()>
        <input type="hidden" name="type" value="人參">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 買 " OnClick=check()>
        <input type="submit" name="submit" value=" 賣 " OnClick=check()>
        <input type="hidden" name="type" value="珠寶">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
<form method="post" action="jys.asp?action=true">
    <td width="25%"> 
        <input type="text" name="num" size="4">
        <input type="submit" name="submit" value=" 買 " OnClick=check()>
        <input type="submit" name="submit" value=" 賣 " OnClick=check()>
        <input type="hidden" name="type" value="瓷器">
        <input type="hidden" name="port" value="<%=port%>">
      </td>
</form>
  </tr>
  <%
Rs.movenext
Loop
Rs.close : set Rs=nothing
Conn.close : set Conn=nothing
end if
%> 
</table>
</body>
</html>