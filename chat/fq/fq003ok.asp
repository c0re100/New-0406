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
Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ϥΦ�w�l�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O FROM �Τ� WHERE id="&info(9),conn
if rs("�k�O")<1000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o1000�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
rs.close
rs.open "select ����,��O FROM �Τ� WHERE �m�W='"&to1&"'",conn
tl=rs("��O")
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='��w�l' and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A����w�l�ܡH');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="& id &" and �֦���="&info(9)
conn.execute "update �Τ� set �k�O=�k�O-2000 where  id="&info(9)
conn.execute "update �Τ� set �k�O=0,��O=��O-2000,���O=���O-2000 where  �m�W='"&to1&"'"
if tl<0  then
conn.execute "update �Τ� set ���H��=���H��+1 where id="&info(9)
conn.execute "update �Τ� set ���A='��' where �m�W='"&to1&"'"
xuedizi="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�ߥX�k����w�l�A<bgsound src=wav/Bombs020.wav loop=1>��w�l�n�b�F<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�Y�W�A<font color=FFC01F>["&to1&"]</font>�Q��w�l�����k�O���L�A��O����2000�A���O���h2000�A���s���1000�A��¼¼�F..."
call boot(to1)
else
xuedizi="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�ߥX�k��<bgsound src=wav/Bombs020.wav loop=1>��w�l�A��w�l�n�b�F<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>�Y�W�A<font color=FFC01F>["&to1&"]</font>�Q��w�l�����k�O���L�A��O����2000�A���O���h2000�A���s���1000�A�S���w�g��R�j�F..."
end if
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
sd(199)="<font color=FFC01F>�i��w�l�j"& xuedizi &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
'��H
sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
%>