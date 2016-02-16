<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" or info(5)="" then Response.Redirect "../error.asp?id=440"%>
<!--#include file="dadata.asp"-->
<%
nickname=trim(info(0))
mypai=trim(info(5))
sql="select 等級 from 用戶 where id="&info(9)
set rs=conn.execute(sql)
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "你不是江湖中人，不能進參加宴會"
response.end
else
dj=rs("等級")
if dj<2 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('你還是江湖小輩，不能參加宴會');window.close();</script>"	
response.end
else

id=request("id")
sql="select 宴會名,擁有者,內力,體力,金幣,類型,數量 from 宴會列表 where ID=" & id & " and 門派='所有' and 參加者='所有'"
Set Rs=connt.Execute(sql)
if rs.bof or rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
connt.close
Response.Write "<script language=javascript>alert('本席不是為你定的，不要亂來！這樣不好玩:P');window.close();</script>"	
response.end
else
yhming=rs("宴會名")
yyou=rs("擁有者")
nl=rs("內力")
tl=rs("體力")
jb=rs("金幣")
lx=rs("類型")
ls=rs("數量")
if ls<1 then
sql1="delete * from 宴會列表  where ID=" & id
connt.execute sql1
sql1="delete * from 宴會者 where 宴會名=" & id
connt.execute sql1
connt.close
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('你來晚，宴會已經結束!');history.back();</script>"	
response.end
else

sql="select ID from 宴會者 where 參加者='" & nickname & "'  and 宴會名=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
sql2="insert into 宴會者(參加者,宴會名) values ('" & nickname & "' , " & id & " )"
connt.execute sql2
sql1="update 用戶 set 內力=內力+"&nl&",體力=體力+"&tl&",金幣=金幣+1 where 姓名='"&nickname&"' "
conn.execute sql1
sql1="update 宴會列表 set 數量=數量-1 where ID=" & id
connt.execute sql1
conn.close
connt.close
set rs=nothing
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【小道消息】"&nickname&"參加了"&yyou&"在江湖大酒店"&yhming&"廳舉行的※"&lx&"※宴會，增長了不少的體力和內力。</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<%else
connt.close
set connt=nothing
	
Response.Write "<script language=javascript>alert('你已參加過這個宴會了，怎麼還來啊!');history.back();</script>"
response.end
end if%>
<%end if%>
<%end if%>
<%end if%>
<%end if%>
<html>
<title>參加宴會成功!</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
input {border: 1px solid; font-size: 9pt; font-family: "新細明體"; border-color: #000000 solid}
-->
</style>

<body bgcolor="#660000">
<p>&nbsp;</p>
<table width="90%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#996633" align="center"><font color="#FFFFFF">參加酒宴成功</font></td>
</tr>
<tr bgcolor="#66FF66">
<td align="center" height="157">
<p><font> <img src="jd/ka1.gif" width="204" height="251"></font></p>
</td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20" class="unnamed1"><b><font color="#FF3333">你經過了一番狼吞虎咽，喝的上面一個模樣，增加內力 <%=nl%>，體力 <%=tl%>,領到禮金？？？個金幣</font></b></td>
</tr>
<tr bgcolor="#0033CC">
<td align="center" height="20"><b><font color="#FF3333"><input  onClick="javascript:window.document.location.href='jd.asp'" value="返 回" type=button name="back">  <INPUT onclick=window.close() type=button value=關閉>  </font></b></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>