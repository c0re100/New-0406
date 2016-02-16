<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10 then Response.Redirect "../error.asp?id=461"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute("update 用戶 set 道德=-3000 where 姓名='"&request("tjname")&"'")
conn.execute("update 通緝令 set 站長審批='批準',批準時間=now(),批準人員='"& info(0) &"' where 通緝人犯='"& Request("tjname") &"'")

tj="<font color=red>站長[" & info(0) & "]成功批準了通緝人犯[" & Request("tjname") & "]的請求，各位大俠以後見到["&Request("tjname")&"]請不要手軟，要為江湖鏟除公害（通輯人犯無法保護）！</font>"

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
sd(199)="<font color=FFD7F4>【通緝批準】</font>"&tj
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%>
<script language=vbscript> 
MsgBox "完成：已經批準該通緝令！" 
location.href = "javascript:history.back()" 
</script>