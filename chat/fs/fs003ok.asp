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
rs.open "select �k�O,�Ȩ�,����,�D�w,�D�w�[ FROM �Τ� WHERE id="&info(9),conn
if rs("�Ȩ�")<1000000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���W�S��100�U�A���a100�U�b���W�A���a�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if rs("�k�O")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o1000�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-1000,�D�w=�D�w+10000,�Ȩ�=�Ȩ�-1000000 where id="&info(9)
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000 where �m�W='"&to1&"'"
rs.close
rs.open "select ����,�D�w,�D�w�[ FROM �Τ� WHERE id="&info(9),conn
ddj=(rs("����")*1440+2200)+rs("�D�w�[")
if rs("�D�w")>=ddj then
conn.execute "update �Τ� set �D�w=�D�w+"&ddj&" where id="&info(9)
end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
bushishu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�B�Ϊk�O�@�����I�ѤU�o��<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�Ȩ�<font color=red>100�U</font>��A�ۤv�D�w�W�[<font color=red>10000</font>�I." 

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
sd(199)="<font color=CFE0B0>�i���I�N�j"& bushishu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>