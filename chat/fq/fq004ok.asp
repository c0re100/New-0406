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
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ϥηm�T�O�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O FROM �Τ� WHERE id="&info(9),conn
if rs("�k�O")<10000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o10000�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,�s�� FROM �Τ� WHERE �m�W='"&to1&"'",conn
money=int(rs("�s��")/5)
randomize timer
kaa=int(rnd*99)+1
if money>=1000000 then
randomize timer
moneyc=1000000
moneyc=killer+kaa
end if
rs.open "select id FROM ���~ WHERE ���~�W='�m�T�O' and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A���m�T�O�ܡH');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="& id &" and �֦���="&info(9)
conn.execute "update �Τ� set �k�O=�k�O-10000 where id="&info(9)
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&moneyc&",�D�w=�D�w-1000 WHERE  id="&info(9)
conn.execute "update �Τ� set �s��=�s��-"&money&" where �m�W='"&to1&"'"
qiangjielin="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>���X�k��<bgsound src=wav/Phant012.wav loop=1>�m�T�O,���<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�j�q�@�n:�m�T  ,�q<font color=FFC01F>["&to1&"]</font>���m�o�s��<font color=FFC01F>["&moneyc&"]</font>��,<font color=FFC01F>["&info(0)&"]</font>���Ӫk�O<font color=FFC01F>3000</font>�I" 
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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=FFC01F>�i�m�T�O�j"&qiangjielin&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>