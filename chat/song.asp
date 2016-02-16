<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
info(0)=info(0)
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(useronlinename," "&info(0)&" ")=0 then Response.Redirect "chaterr.asp?id=001"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj1=n & "-" & y & "-" & r
sj2=s & ":" & f & ":" & m
sj3=sj1 & " " & sj2
if chatbgcolor="" then chatbgcolor="008888"%>
<!--#INCLUDE FILE="midsound.asp" -->
<html>
<head>
<title>點歌</title>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<script language="javaScript">
function send(){sendurl="songsend.asp?songname=" + document.forms[0].song.options[document.forms[0].song.selectedIndex].value;window.location.href=sendurl;}
</script>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" text="#FFFFFF" topmargin="0" background=bk.jpg>
<div align=center>
  <table width="135" border="0">
    <tr>
      <td> 
        <div align=center><font color="#FFFFFF" style="font-size:12pt">點歌系統</font></div>
      </td>
    </tr>
  </table>
  
</div>
<table border="0" width="154" align="center" height="267">
  <form method="post" action="songplay.asp" name="" target="ps">
<tr>
<td>
<div>
<p><a href="javascript:history.go(0)" title="顯示最新的曲目列表"><font color="#FFFFFF">刷新:<%=sj1%></font></a></p>
</div>
<p> <font color="#FFFF00">選擇曲目：</font><br>
<select name="song" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<%For t=1 to 85%>
<option value="<%=t%>"><%=mid(t)%></option>
<%next%>
erase mid

</select>
</p>
<p><font color="#FFFF00">播放方式：</font><br>
<input type="radio" name="loopok" value="1" checked>
隻聽一遍<br>
<input type="radio" name="loopok" value="infinite">
百聽不厭</p>
<table border="0" cellpadding="4">
<tr>
<td>
<input type="submit" name="play" value="播放">
</td>
<td align="right">
<input type="submit" name="st" value="停止">
</td>
</tr>
<tr>
<td colspan="2">
<input type="button" value="好歌贈好友" onClick="Javascript:send()" name="button">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
<Script Language=Javascript>
parent.m.location.href='about:blank';
</Script>
</body>
</html> 