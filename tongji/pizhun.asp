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
conn.execute("update �Τ� set �D�w=-3000 where �m�W='"&request("tjname")&"'")
conn.execute("update �q�r�O set �����f��='���',��Ǯɶ�=now(),��ǤH��='"& info(0) &"' where �q�r�H��='"& Request("tjname") &"'")

tj="<font color=red>����[" & info(0) & "]���\��ǤF�q�r�H��[" & Request("tjname") & "]���ШD�A�U��j�L�H�ᨣ��["&Request("tjname")&"]�Ф��n��n�A�n�������갣���`�]�q��H�ǵL�k�O�@�^�I</font>"

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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�q�r��ǡj</font>"&tj
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%>
<script language=vbscript> 
MsgBox "�����G�w�g��Ǹӳq�r�O�I" 
location.href = "javascript:history.back()" 
</script>