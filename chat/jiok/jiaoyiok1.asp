<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
grade=info(2)
myid=info(9)
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if application("jiutian_jiaoyi")="" then
Response.Write "<script Language=Javascript>alert('�藍�_�A�{�b�S�����A�Ϊ̩��w�g����');window.close();</script>"
response.end
end if

sj=DateDiff("s",Application("jiutian_shijian"),now())
if sj<3 then
s=3-sj
Response.Write "<script language=JavaScript>{alert('���ܡG�����b���B�z�ƾڡA�бz���W["&s&"����]�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
Application.Lock
Application("jiutian_shijian")=now()
Application.UnLock


jian=split(application("jiutian_jiaoyi"),"|")
from1=jian(0)
wupin=jian(3)
jiage=int(jian(2))
jia=int(jiage*0.9)
sl=int(jian(4))
danwei=jian(5)
froo=jian(6)
if info(0)=froo then
Response.Write "<script Language=Javascript>alert('�藍�_�A�ۤv����R�ۤv���F��');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select �m�W from �Τ� where id="&from1
rs.open sql,conn,1,1
nickk=rs("�m�W")
rs.close
sql="SELECT ���~�W,�˳� FROM ���~ WHERE (����='�t��' or ����='�ī~' or ����='�k��' or ����='�k�_')  and �֦���="&from1
rs.open sql,conn,1,1
wqm=rs("���~�W")
my_wu=wupin
if my_wu=wqm and rs("�˳�")=true then
Response.Write "<script Language=Javascript>alert('�ʶR�����\�A��誫�~�˳Ƶ۩O');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if

rs.close


sql="select �Ȩ� from �Τ� where id="&info(9)
rs.open sql,conn,1,1

jinyan=int(rs("�Ȩ�"))

if jiage>jinyan then
Response.Write "<script Language=Javascript>alert('�ʶR�����\�A�A���W�S���o��h���Ȩ�');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if

rs.close
sql="select * from ���~ where ���~�W='" & wupin & "' and �֦���="&from1
rs.open sql,conn,1,1
if rs.eof or rs.bof then
Response.Write "<script Language=Javascript>alert('�ʶR�����\�A���S���o�˪����~');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
wu=rs("���~�W")
yin=rs("�Ȩ�")
dj=rs("����")
lx=rs("����")
gj=rs("����")
fy=rs("���s")
nl=rs("���O")
tl=rs("��O")
say=rs("����")
sm=rs("sm")
rs.close
sql="select �ƶq from ���~ where ���~�W='" & wupin & "' and �֦���="&info(9)
rs.open sql,conn,1,1
if rs.eof or rs.bof then
conn.execute"insert into ���~(���~�W,�֦���,����,����,���O,��O,����,���s,�ƶq,�Ȩ�,����,sm,�|��) values ('"&wu&"',"&info(9)&",'"&lx&"','"&say&"','"&nl&"','"&tl&"','"&gj&"','"&fy&"',"&sl&","&yin&","&dj&","&sm&","&aaao&")"
conn.execute"update ���~ set �ƶq=�ƶq-'"&sl&"'  where ���~�W='" & wupin & "' and �֦���="&from1
if danwei="�Ȩ�" then
conn.execute"update �Τ� set �Ȩ�=�Ȩ�+'" & jia & "' where id="&from1
conn.execute"update �Τ� set �Ȩ�=�Ȩ�-'" & jiage & "' where id="&info(9)
else
conn.execute"update �Τ� set �Ȩ�=�Ȩ�+'" & jia & "' where id="&from1
conn.execute"update �Τ� set �Ȩ�=�Ȩ�-'" & jiage & "' where id="&info(9)

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
msg="["&nickname&"]��F<font color=red>"&jiage&"</font>�Ȩ�A�q["&nickk&"]���R�F<font color=FFFDAF>"&wupin&"</font>,������\�A�x���q������10%���|�C"
call jiaoyi()
Response.Write "<script Language=Javascript>alert('������\');window.close();</script>"
response.end
end if

shu=rs("�ƶq")
if shu>20 then

Response.Write "<script Language=Javascript>alert('�A���W�����~�Ӧh�A����A�R!');window.close();</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
else
conn.execute"update ���~ set �ƶq=�ƶq+'"&sl&"'  where ���~�W='" & wupin & "' and �֦���="&info(9)
conn.execute"update ���~ set �ƶq=�ƶq-'"&sl&"'  where ���~�W='" & wupin & "' and �֦���="&from1
if danwei="�Ȩ�" then
conn.execute"update �Τ� set �Ȩ�=�Ȩ�+'" & jia & "' where id="&from1
conn.execute"update �Τ� set �Ȩ�=�Ȩ�-'" & jiage & "' where id="&info(9)
else
conn.execute"update �Τ� set �Ȩ�=�Ȩ�+'" & jia & "' where id="&from1
conn.execute"update �Τ� set �Ȩ�=�Ȩ�-'" & jiage & "' where id="&info(9)

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
msg="["&nickname&"]��F<font color=red>"&jiage&"</font>�Ȩ�A�q["&nickk&"]���R�F<font color=FFFDAF>"&wupin&"</font>,������\�A�x���q������10%���|�C"
call jiaoyi()
Response.Write "<script Language=Javascript>alert('������\');window.close();</script>"
response.end
end if

sub jiaoyi()
application("jiutian_jiaoyi")=""
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
sd(199)="<font size=2><font color=FFFDAF>�i��榨��j" &msg &"</font></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
%>