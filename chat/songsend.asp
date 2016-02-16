<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
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
songname=Request.QueryString("songname")
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
if chatbgcolor="" then chatbgcolor="008888"%>
<!--#INCLUDE FILE="midsound.asp" -->
<html>
<head>
<title>好歌贈好友</title>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<link rel="stylesheet" href="readonly/style.css">
<script language="JavaScript">
function gozj(){url="songsend.asp?zj="+document.forms[0].songzj.options[document.forms[0].songzj.selectedIndex].value;window.location.href=url;}
function check(){document.forms[0].send.disabled=1;return true;}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" text="#FFFFFF" topmargin="0">
<div align="center"><font color="#FFFF00" style="font-size:12pt">好歌贈好友</font></div>
<hr size=1 color=FFFF00><br>
<table border="0" width="154" align="center" height="267">
  <form method="post" action="songsendok.asp" onsubmit='return(check());'>
<tr>
<td>
<p><font color="#FFFF00">選擇曲目：</font><br>
<select name="song" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<%For t=1 to 85%>
<option value="<%=t%>"><%=mid(t)%></option>
<%next%>
erase mid
</select>
</p>
<p><font color="#FFFF00">選擇好友：</font><br>
<select name="towho" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'><%for i=0 to x%>
<option value="<%=online(i)%>"<%if CStr(online(i))=CStr(info(0)) then response.write " selected"%>><%=online(i)%></option>
<%next%></select>
</p>
<p> <font color="#FFFF00">祝福話語：</font><br>
<input type="text" name="infofo" size="16" style='font-size:12px;color:#FF6633;background-color:#EEFFEE' maxlength="100" value="謹以《%%》表達我的心意  "></p>
<p>注：%%代表歌名
</p>
<table border="0">
<tr>
<td>
<input type="submit" name="send" value="發送">
<input type="button" name="abort" value="放棄" onClick="javascript:history.go(-1)">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</body>
</html> 