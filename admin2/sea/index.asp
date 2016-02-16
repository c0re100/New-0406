<%
Response.Expires=0
Response.Buffer=True
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
action=Request.Querystring("action")
if action="search" then
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Randomize timer
r=int(rnd*10)
Set Conn1=Server.CreateObject("ADODB.Connection")
Connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("global.asa")
Conn1.Open Connstr1
set rs=conn1.execute("select 次數,時間 from 航海物品 where 擁有者='"& name &"'")
if rs.eof and rs.bof then
  Conn1.Execute("insert into 航海物品(擁有者) values('"& name &"')")
times=8
else
times=rs("次數")
  if now()-rs("時間")>1 then
    conn1.execute("update 航海物品 set 次數=次數+8,時間=now() where 擁有者='"& name &"'")
    times=times+8
  end if
end if
        if times<1 then
            mess="你的行動力沒了！！！"
        else
            conn1.execute("update 航海物品 set 次數=次數-1 where 擁有者='"& name &"'")
            if r=5 then
              conn.execute("update 用戶 set 銀兩=銀兩+100000 where 姓名='"& name &"'")
              mess="發現了寶藏10萬兩！！！"
            else
              mess="一無所獲！！！"
            end if
        end if
set rs=nothing
conn1.close : set conn1=nothing
conn.close : set conn=nothing
response.write "<script language='javascript'>alert('"& mess &"');window.close();</script>"
else
%>
<html>
<head>
<title>航海時代</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "新細明體"; font-size: 12px}
input {  font-family: "新細明體"; font-size: 12px; border: thin #333333 dotted; color: #660000; background-color: #FFFFCC}
-->
</style>
<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body bgcolor="#FFFFFF" text="#660000" background="../../images/8.jpg">
<br>
<br>
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr align="center"> 
    <td height="22"><big><strong>江湖航海時代</strong></big><br>
      <br>
      (每天可做<font face="Verdana">8</font>筆傻瓜似的生意) 
      <input type="button" Onclick="MM_openBrWindow('see.asp','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="商品一纜"> 
      此插件測試中，發現錯誤請告訴我！！！ <br> 
    </td> 
  </tr> 
</table> 
<table width="90%" border="1" cellspacing="4" cellpadding="4" align="center">
  <tr align="center" bgcolor="#660000">  
    <td width="25%"><font color="#FFFFFF">大連港</font></td> 
    <td width="25%"><font color="#FFFFFF">杭州港</font></td> 
    <td width="25%"><font color="#FFFFFF">寧波港</font></td> 
    <td width="25%"><font color="#FFFFFF">泉洲港</font></td> 
  </tr> 
  <tr align="center">  
    <td width="25%">  
      <input type="button" Onclick="MM_openBrWindow('jys.asp?port=大連','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所"> 
      <input type="button" Onclick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城門古跡"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=杭州','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所" name="button"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城門古跡" name="button"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=寧波','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所" name="button2"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城門古跡" name="button2"> 
    </td> 
    <td width="25%">  
      <input type="button" onClick="MM_openBrWindow('jys.asp?port=泉洲','','scrollbars=yes,width=600,height=200,top=100,left=50')" value="交易所" name="button3"> 
      <input type="button" onClick="MM_openBrWindow('index.asp?action=search','','scrollbars=yes,width=300,height=300')" value="城門古跡" name="button3"> 
    </td> 
  </tr> 
</table> 
<table width="90%" border="0" cellspacing="0" cellpadding="0" align="center"> 
  <tr align="center">  
    <td height="22"><br>
      靈劍江湖 <br>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>
