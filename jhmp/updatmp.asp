<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
comment=request.form("comment")
shihe=trim(request.form("shihe"))
zhangmen=trim(request.form("zhangmen"))
s=Request.QueryString("subject")
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
if chuser(zhangmen) then Response.Redirect "../error.asp?id=120"
if shihe<>"�k" and shihe<>"�k" and shihe<>"�Ҧ�" then Response.Redirect "../error.asp?id=442"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if zhangmen="�L" then
rs.open "select * FROM ���� WHERE ����='"&s&"' and �x��='"&info(0)&"' ",conn
if rs.EOF or rs.BOF then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('��p�A���O�x������ާ@�I');history.back();</script>"
Response.End
else
conn.execute "Update ���� Set �A�X='" & shihe & " ',²��='" & comment & " ' Where ����='" & s & "'"
end if
else
rs.open "select id FROM �Τ� WHERE ����='"&s&"' and �m�W='"&zhangmen&"' ",conn
if rs.EOF or rs.BOF then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('�n�Ǧ쪺�H���O�A���̤l�I');history.back();</script>"
Response.End
else
conn.execute "Update ���� Set �x��='"&zhangmen&"' Where ����='" & s & "'"
conn.execute "Update �Τ� Set ����='" & s & "',����='�x��',grade=5 Where �m�W='"&zhangmen&"'"
conn.execute "Update �Τ� Set ����='" & s & "',����='�̤l',grade=1 Where �m�W='"&info(0)&"'"
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
sd(195)=info(0)
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�t�Ρj[" & s & "]���x��:"& info(0) &"</font><font color=FFD7F4>��x���_�y�ǵ��F "&zhangmen&" �I</font></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end if
end if
Response.Redirect "../ok.asp?id=701"
%> 