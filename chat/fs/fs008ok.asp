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
	Response.Write "<script Language=Javascript>alert('�ҭn�I�Ϫ��H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ϥΦ��k�N�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
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
if rs("�k�O")<1000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o1000�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM �Τ� WHERE id="&info(9),conn
conn.execute "update �Τ� set �k�O=�k�O-1000,�ާ@�ɶ�=now() where id="&info(9)
rs.close
rs.open "select ����,��O�[ FROM �Τ� WHERE �m�W='"&to1&"'",conn
tlj=(rs("����")*1500+5260)+rs("��O�[")
conn.execute "update �Τ� set ��O=" & tlj & " where �m�W='"&to1&"'"
jijiushu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>���ݵ�<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�N�֭n�N����A�N������H�u�I�l�I�ϡA<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>��O���t�ɥR�ܷ����P�E���s�s�V<font color=FFC01F>["&info(0)&"]</font>��ܷP��." 

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
sd(199)="<font color=CFE0B0>�i��ϳN�j"& jijiushu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>