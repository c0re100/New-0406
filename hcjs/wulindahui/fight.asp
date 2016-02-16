<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data.asp"-->
<%
my=info(0)
types=Request.QueryString("typename")
id=Request.QueryString("id")
Select Case id
Case "1"
  classes="狀元"
Case "2" 
   classes="榜眼"
Case "3"
   classes="探花"
End Select
Select Case types
Case "gold"
  typeses="金榜"
Case "silver" 
   typeses="銀榜"
Case "copper"
   typeses="銅榜"
End Select
%>
<%
sql="select 武功,防御,攻擊,內力,身份,體力,性別 from 用戶 where 姓名='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	Response.Redirect "../../error.asp?id=454"
	conn.close
	response.end
end if
set rs1=server.createobject("adodb.recordset")
sql1="select 攻擊,防御,內力,武功 from "&types&" where id="&id
set rs1=conn1.execute(sql1)
if rs("武功")>10000 and types="copper" then
%>
<script language=vbscript>  
  MsgBox "你如此厲害的武功已經不屑於來揭銅榜了。"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("武功")<=10000 and types="silver" then
%>
<script language=vbscript>  
  MsgBox "你如此差的武功還想來揭銀榜？回去再練練吧！"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("武功")>20000 and types="silver" then
%>
<script language=vbscript>  
  MsgBox "你如此厲害的武功已經不屑於來揭銀榜了。"  
  location.href = "meeting.asp"  
</script>
<%
rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("武功")<20000 and types="gold" then
%>
<script language=vbscript>  
  MsgBox "你如此差的武功還想來揭金榜？回去再練練吧！"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
set rs2=server.createobject("adodb.recordset")
sql2="select id from "&types&" where  姓名='" & my & "'"
set rs2=conn1.execute(sql2)
if not rs2.eof then
if (rs2("id")<3 and id=3) or (rs2("id")<2 and id=2) then
%>
<script language=vbscript>
MsgBox "你已經取得比這個更高的功名了，太貪心了吧！"
location.href = "meeting.asp"
</script>
<%
rs.close
rs1.close
rs2.close
conn.close
conn1.close
response.end
end if
if (rs2("id")=3 and id=3) or (rs2("id")=2 and id=2) or (rs2("id")=1 and id=1) then
%>
<script language=vbscript>
MsgBox "你已經取得這個功名了，想和自己過不去嗎?"
location.href = "meeting.asp"
</script>
<%
rs.close
rs1.close
rs2.close
conn.close
conn1.close
response.end
end if
end if
if rs("防御")>rs1("防御") and rs("攻擊")+rs("內力")>rs1("攻擊")+rs1("內力") then
 wugong=rs("武功")-rs1("武功")
 defence=rs("防御")-rs1("防御")
 force=rs("攻擊")-rs1("攻擊")
 neili=rs("內力")-rs1("內力")
 sql1="update "&types&" set 姓名='"& my & "',性別='"&rs("性別")&"',門派='"&rs("門派")&"',身份='"&rs("身份")&"',武功="&rs("武功")&",內力="&rs("內力")&",體力="&rs("體力")&",攻擊="&rs("攻擊")&",防御="&rs("防御")&" where id="&id&""
conn1.execute(sql1)
sql="update 用戶 set 內力=內力+500,體力=體力-50,allvalue=allvalue+100, 銀兩=銀兩+500000 where 姓名='"&my&"'"
conn.execute(sql)
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
sd(199)="<font color=FFFDAF>【公告】</font><font color=#DFF0CF>恭喜<font color=#DFF0CF>"& my &"</font>揭榜成功，登上了<font color=#DFF0CF>"& typeses & classes &"</font>的寶座!站長獎勵，內力5000，經驗100點，銀子500000！！"&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language=vbscript>  
  MsgBox "恭喜你，你揭榜成功！"  
  location.href = "meeting.asp"  
</script>
<%
else
sql="update 用戶 set  內力=內力-50,體力=體力-50,銀兩=銀兩-500 where 姓名='"&my&"'"
conn.execute(sql)
%>
<script language=vbscript>  
  MsgBox "對不起！你揭榜失敗！結果被打得鼻青臉腫的。"  
  location.href = "meeting.asp"
</script>
<%
end if
%>