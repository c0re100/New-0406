<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�ҧ������H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�k�O,����,����,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
dj=rs("����")
shenfen=rs("����")
fla=rs("�k�O")
if fla<10000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o100�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj>100 or shenfen="�x��" or rs("����")="�x��" then
Response.Write "<script language=JavaScript>{alert('�A�O�x���کε���100�ŤF�o��F�`�٥μa�A�N�r�I');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select �ʧO,�Ȩ� FROM �Τ� WHERE �m�W='"&to1&"'",conn
xb=rs("�ʧO")
money=int(rs("�Ȩ�")/4)
randomize timer
kaa=int(rnd*99)+1
if money>=1000000 then
randomize timer
moneyc=1000000
moneyc=killer+kaa
end if
if xb="�k" then
Response.Write "<script language=JavaScript>{alert('���a�A�N�u��k�Ĥl�A�ΡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&moneyc&",�k�O=�k�O-10000,�ާ@�ɶ�=now() WHERE  id="&info(9)
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&money&" where �m�W='"&to1&"'"

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
meihuoren="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>��<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�I�i<bgsound src=wav/phant030a.wav loop=1>�y�b�H�����A�N�A<font color=FFC01F>"&to1&"</font>�Q<img src='img/007.gif'><font color=red>["&info(0)&"]</font>�ɰ�ɫ����e���Ұg�b�A���W���Ȥl������ı�֤F�|�����@,�@�p<font color=red>"&moneyc&"</font>��:)..." 
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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=CFE0B0>�i�a�A�N�j"& meihuoren &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>