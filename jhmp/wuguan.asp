<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>靈劍江湖練武</title>
<body leftmargin="0" topmargin="0" bgcolor="#003399" background="../linjianww/0064.jpg">
<table border="0" cellspacing="0" cellpadding="0" width="97" align="center">
<tr>
<td height="150" valign="top">
      <div align="center"><font color="#000000"><b><font color=blue><%=info(0)%></font>歡迎光臨習武場</b></font></div>
<form method=POST action='wuguanok.asp'>
<table width="300" align="center">
<tr>
<td>
<tr>
            <td align=center><font color=blue>師傅:<%=info(7)%></font>
            <%if trim(info(7))<>"無" then%>
              <select name=money size=1>
                <option value="1000" selected> 打馬步（一千兩）</option>
                <option value="10000">少林硬功（一萬兩）</option>
                <option value="100000">峨眉心法（十萬兩）</option>
                <option value="1000000">逍遙神劍（一百萬兩）</option>
                <option value="2000000">苗家蠱毒（二百萬兩）</option>
                <option value="10000000">浩然正氣（一千萬兩）</option>
              </select>
              <%end if%>
</td>
</tr>
<tr>
<td  align=center>
            <%if trim(info(7))<>"無" then%>
              <input type=submit value=開始練功 name="submit">
              <%end if%>
</td>
</tr>
<tr>
<td valign="top" height="8" >
<div align="center"><br>
<br>
                操作簡介</div>
</td>
</tr>
<tr>
            <td valign="top" > 
              <p><br>
            <%if trim(info(7))="無" then%><div align="center">
            <font color=red>沒有師傅不能操作！</font><BR></div>
            <%end if%>師傅領你去練功，選擇你適合自己的心法(價錢是不同的喲！)纔可以修天到至高無限的武功！</p>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
</body>
</html>



