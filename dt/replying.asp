<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
mycorp=info(5)
if username=Application("ljjh_askuser") then response.redirect "../error.asp?id=476"
if mycorp="�x��" then response.redirect "../error.asp?id=477"
reply=trim(Request.form("reply"))
reply=server.HTMLEncode(reply)
if trim(Request.form("reply"))="" then Response.Redirect "../error.asp?id=475"
if reply=Application("ljjh_reply") then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �Ȩ� FROM �Τ� WHERE id="&info(9)
rs.open sql,conn,1,3
rs("�Ȩ�")=rs("�Ȩ�")+Application("ljjh_asksilver")
rs.Update
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="�i�� �D�j<font color=FFFDAF>" & Application("ljjh_ask") & "�H</font>���T���׬O�G<font color=red>[ " & Application("ljjh_reply") & " ] </font> "& username &" �]����o�F<font color=FFFDAF>"& Application("ljjh_asksilver") &"</font>��Ȥl�����y!"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application("ljjh_talkarr")=newtalkarr
Application("ljjh_ask")=""
Application("ljjh_reply")=""
Application("ljjh_askuser")=""
Application.ULock
else
response.redirect "../error.asp?id=479"
response.end
end if
%>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>

