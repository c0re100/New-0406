<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<1 then
ss=1-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('請等上"&ss&"分再來碰運氣吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<head>
<title>江湖抽獎</title>
<meta http-equiv="refresh" content="10">
<style></style>
</head>
<BODY oncontextmenu=self.event.returnValue=false BGCOLOR="#fefefe" text="#FFFFFF" > 
<DIV ID="Layer1" STYLE="position:absolute; left:46px; top:139px; width:192px; height:165px; z-index:3">
hu4<%
if Session("choujiang")=true then
	Session("choujiang")=now()
end if 
abc=DateDiff("s",Session("choujiang"),now())
if abc>=200 then
%>
<div align="center">
<script language=vbscript>
location.href = "pao.asp"
</script>
<%
response.end
end if
if abc<=60 then
%>
<IMG SRC="CHOUJIANG.gif"><br>
<font color="#0000FF">請等吧，抽獎是須時的！</font>
<%
else
%>
<a href="choujiangok.ASP"><IMG SRC="choujiangok.jpg" border="0" width="213" height="185"></a>
<font color="#0000FF">結果公佈了，快到公怖板看看！</font>
<%
end if
%>
</div>
</DIV>
<DIV ID="Layer2" STYLE="position:absolute; left:11px; top:13px; width:252px; height:102px; z-index:1"><IMG SRC="CHOU.jpg"></DIV>
<span id="liveclock" style="position: absolute; left: 328px; top: 275px; width: 76px; height: 23px">
<div align="center"><font color=red>抽獎進行時間:<br>
<%=abc%>秒</font> </div>
</span>
</body>
</html>