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
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�ϥΰg��j�k�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,����,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<9 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=9-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("�k�O")<3000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o3000�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select �k�O,���� FROM �Τ� WHERE �m�W='"&to1&"'",conn
falit=int(rs("�k�O")/2)
if Application("ljjh_dubozhang2")=to1 then
Response.Write "<script language=JavaScript>{alert('���{�b�O�k�O�䧽�����a���i�H�l��k�O�I');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-2000+"&falit&",�ާ@�ɶ�=now() where id="&info(9)
conn.execute "update �Τ� set �k�O=�k�O-"&falit&" where �m�W='"&to1&"'"
mihunfa="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>��<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�I�i<bgsound src=wav/fx40a.wav loop=1>�g��j�k�A<font color=FFC01F>["&to1&"]</font>�k�O�v���Q<font color=red>["&info(0)&"]</font>�ҧl����"&falit&"�I�A<font color=FFC01F>["&to1&"]</font>���ʦa�b���̯��ۡA�I..." 
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
sd(199)="<font color=CFE0B0>�i�g��N�j"& mihunfa &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>