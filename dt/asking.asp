<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)>=6  then
ask=trim(Request.form("ask"))
reply=trim(Request.form("reply"))
ask=lcase(server.HTMLEncode(ask))
reply=lcase(server.HTMLEncode(reply))
silver=Trim(Request.Form("silver"))
if not isnumeric(silver) then Response.Redirect "../error.asp?id=475"
if trim(Request.form("ask"))="" then Response.Redirect "../error.asp?id=475"
if trim(Request.form("reply"))="" then Response.Redirect "../error.asp?id=475"
silver=clng(silver)
if silver<1 then
silver=1
elseif silver>9999999 then
silver=9999999
end if
ask=replace(ask,"'","")
ask=replace(ask,"or","")
ask=replace(ask,"\","")
ask=replace(ask,"www","")
ask=replace(ask,"xajh","")
ask=replace(ask,"net","")
reply=replace(reply,"'","")
reply=replace(reply,"or","")
reply=replace(reply,"\","")
reply=replace(reply,"www","")
reply=replace(reply,"xajh","")
reply=replace(reply,"net","")
Application.Lock
Application("ljjh_ask")=ask
Application("ljjh_reply")=reply
Application("ljjh_askuser")=info(0)
Application("ljjh_asksilver")=silver
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next

sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="�i�X �D�j<font color=FFFDAF>" & Application("ljjh_ask") & "�H </font>���T���׬O����H[���ݤH]:<font color=red> " & Application("ljjh_askuser")&" </font><font color=FFFDAF>���y�G"&Application("ljjh_asksilver")&"��!</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>
<%
else
response.redirect "../error.asp?id=425"
end if
%>
