<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(0)<>"�����`��" then Response.Redirect "../error.asp?id=439"
a=request.form("a") '���ID
b=request.form("b") '���~�W
c=request.form("c") '���~�ƶq
d=request.form("d") '����
e=trim(request.form("e")) '����
f=request.form("f") '���O
g=request.form("g") '��O
h=request.form("h") '����,
i=request.form("i") '���m
j=request.form("j") '�Ȩ�
k=request.form("k") '��T��
m=request.form("m") '����
n=request.form("n") '�|��
l=request.form("l") 'sm
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

rs.open "SELECT id FROM �Τ� where �m�W='"&a&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�S���ӥΤ�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
idd=rs("id")
rs.close
rs.open "select id from ���~ where ���~�W='"&b&"' and �֦���="&idd,conn
if rs.eof and rs.bof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,���O,��O,����,���s,�ƶq,�Ȩ�,��T��,����,sm,�|��) values ('"&b&"',"&idd&",'"&d&"','"&e&"',"&f&","&g&","&h&","&i&","&c&","&j&","&k&","&m&","&l&","&n&")"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ܡG�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
id=rs("id")
sql="update ���~ set �ƶq=�ƶq+"&c&" where id="& id
conn.execute(sql)
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
sd(199)="<font color=#ffff00>�i�ʶR�����j</font>"& a&"�V�����`���ʶR�F<font color=#ffffcc><b>"& b &"�@"& c&"��</b></font>...�j�a���ߥL"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>
