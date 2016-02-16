<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<6 then Response.Redirect "../error.asp?id=461" 
if trim(request.form("tjname"))="" or trim(request.form("yuanyin"))="" or trim(Request.Form("fuyan"))="" then %> 
<script language=vbscript> 
MsgBox "錯誤：通緝人犯、通緝原因、請求附言必須填寫！" 
location.href = "javascript:history.back()" 
</script> 
<%end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'rs.open "select 等級 from 用戶 where 姓名='" & request.form("tjname") & "'"
	sql="SELECT * FROM 用戶 WHERE 姓名='" & request.form("tjname") & "'"
	Set Rs=conn.Execute(sql)
	if rs.eof or rs.bof then

	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('你通緝的並非江湖人員，請仔細檢查您輸入的通緝人犯的填寫正確!');history.back();</script>"
	response.end
end if%>
<%
tjname=request.form("tjname")
yuanyin=request.form("yuanyin")
fuyan=request.form("fuyan")
fabu=info(0)
%>
<!--#include file="data1.asp"--><%
'rs.close
'rs.open "select 通緝人犯 from 通緝令 where 通緝人犯='" & request.form("tjname") & "'"
sql="SELECT * FROM 通緝令 WHERE 通緝人犯='" & request.form("tjname") & "'"
	Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
                sql="insert into 通緝令(通緝人犯,請求人員,通緝原因,請求附言,請求時間,站長審批,批準時間,批準人員) values ('" & tjname & "','" & fabu & "','" & yuanyin & "','" & fuyan & "',now(),'待審','未知','未知')"
                        conn.execute sql
tj="<font color=red>聊管[" & fabu & "]成功提交通緝請求，通緝人犯：[" & tjname & "]，通緝原因：[" & yuanyin & "]，狀態：等待站長審批！</font>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
sd(199)="<font color=FFD7F4>【通緝請求】</font>"&tj
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language=vbscript> 
MsgBox "完成：您已成功發布通緝請求，請耐心等待站長審批！" 
location.href = "javascript:history.back()" 
</script>
<%else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript> 
MsgBox "錯誤：該人犯已經在請求名單內，請勿重復填寫！" 
location.href = "javascript:history.back()" 
</script> 

<%end if%>