<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if session("nowinroom")<>1 then
Response.Write "<script Language=Javascript>alert('�A�b�������~�԰϶ܡH');Javascript:window.close();</script>"
	Response.End
end if
pkshijian=Hour(now())
if pkshijian<>21 then
Response.Write "<script Language=Javascript>alert('�ݦn�ɶ��F�A�����Աq��20�G00��21�G00�A�{�b�٦��I');Javascript:window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from �����j�� where �D������='"&info(5)&"' or �Ĺ�����='"&info(5)&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
  response.write "<Script>alert('[�H�a�����j�ԧA�M�ͦX����H�I');Javascript:window.close();</script>"
  response.end
end if
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
onliners=Application("ljjh_chatrs"&session("nowinroom"))
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
if x>0 then
rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
  response.write "<Script>alert('[���ɳW�h�G���঳�@�ӳӧQ�̡A���~:�s�]�@�ӥѸӬ����D�Ҧ��A����U�̤l�Ȩ�1000�U�B����1000�B�k�O1000�I');Javascript:window.close();</script>"
  response.end
end if
name=online(0)
rs.close
	rs.open "select ���� from �Τ� where �m�W='"&name&"'",conn
menpai=rs("����")
rs.close
	rs.open "select �x�� from ���� where ����='"&menpai&"'",conn
zhangmen=rs("�x��")
rs.close
	rs.open "select id from �Τ� where �m�W='"&zhangmen&"'",conn
idd=rs("id")
rs.close
rs.open "select id from ���~ where ���~�W='�s�]' and �֦���=" & idd,conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,����,�|��) values ('�s�]',"&idd&",'���~',0,100,0,0,0,1,10000000,50000,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
 end if
conn.Execute"update �Τ� set �s��=�s��+10000000,����=����+1000,�k�O=�k�O+1000 where ����='" & menpai & "'"
conn.Execute"update ���� set �x��='"&name&"' where ����='" & menpai & "'"
call boot(info(0))
bangpaizhan="����["&info(5)&"]��o�ӧQ!"
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
sd(195)=info(0)
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i�����j�ԡj</font>"&bangpaizhan
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
  response.write "<Script>alert('[���ߧA����o�ӧQ!���w�����D�Э��s�i�J�I]');Javascript:window.close();</script>"
  response.end
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
 