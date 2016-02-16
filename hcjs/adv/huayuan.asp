<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data.asp"-->
<!--#include file="func.asp"-->
<%
sh=request.form("sh")
'sy=request.form("sy")
my=info(0)
if request.form("h")="1" then
sql="select 體力,狀態,操作時間 from 用戶 where 姓名='" & info(0) & "'"
set rs=conn.execute(sql)
if rs("體力")<-1000 or rs("狀態")="死" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../chat/chaterr.asp?id=001"
	response.end
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<30 then
	s=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('警告：請等"& s &"秒再冒險,你想沒命呀！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if rs("體力")<20 then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你疲勞程度已超出範圍，為防不測，還是離開孤島為上！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sql="update 用戶 set 體力=體力-20 where 姓名='" & info(0) & "'"
	conn.execute sql
        rs.close
        set rs=nothing
	conn.close
        set conn=nothing
        message=huayuan(my,sh,sy)
end if
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=#ff0000>消息</font>："&message
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body  bgproperties="fixed" vlink="#000000" bgcolor="#FFFFFF">
<center>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
        <table height="260" align="center" width="201">
          <tr>
            <td height="37"> 
              <div align="center"><font color="#000000"><strong>突發事件</strong></font></div>
            
          <tr>
  <td height="182" valign="top"><%=message%><Br><Br>
  </td>
  </tr>
          <td align=center height="29">&nbsp;
            <div align="right">
                <input type=button value="返 回 孤 島" onclick="location.href='index.asp'">
                <input type=button value="離 開 孤 島" onclick="location.href='../../welcome.asp'">
 </div>
          </td></tr>
</table>
</td></tr>
</table>
</center>
</body>
</html>
