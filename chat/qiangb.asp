<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
if Application("ljjh_qiang")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ӱߤF�A�F��w�g�Q�O�H�m���F...');</script>"
	response.end
end if
Application.Lock
Application("ljjh_qiang")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ���~�W,�Ȩ�,����,���O,��O,����,��T��,����,sm,�|�� FROM ���~ where id=" & ID,conn
wu=rs("���~�W")
yin=rs("�Ȩ�")
lx=rs("����")
nl=rs("���O")
tl=rs("��O")
say=rs("����")
sm=rs("sm")
dj=rs("����")
jg=rs("��T��")
rs.close
'���~
rs.open "select id from ���~ where ���~�W='"&wu&"' and �֦���="& info(9),conn
If Rs.Bof OR Rs.Eof then
	sql="insert into ���~(���~�W,�֦���,����,���O,��O,��T��,�ƶq,�Ȩ�,����,����,sm,�|��) values ('"&wu&"',"&info(9)&",'"&lx&"',"&nl&","&tl&","&jg&",1,"&yin&","&dj&",'"&say&"',"&sm&","&aaao&")"
	conn.execute sql
else
	id1=rs("id")
	sql="update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="&id1
	conn.execute sql
end if
conn.execute "update ���~ set �ƶq=�ƶq-1,�|��="&aaao&" where id="&id
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�m��F�@�˪F��A�֬ݫ̹��I');</script>"
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="��"
sd(199)="<font color=9FDF70>�i�m�r�j"&info(0)&"��i���s������m��@��["&wu&"]�C</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

%> 