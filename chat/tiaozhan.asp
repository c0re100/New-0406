<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
yn=LCase(trim(request.querystring("yn")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W,����,�W���Y�� FROM �Τ� WHERE id="&id,conn
peiou=rs("�m�W")
menpai=rs("����")
old=rs("�W���Y��")
if info(0)<>trim(Application("dantiao")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�Q�@����A�S�H�V�A�U�D�ԮѰڡI');}</script>"
	Response.End
end if
if yn=1 then
	Response.Write "<script language=JavaScript>{alert('�ڵ����A���n�D�F�I');}</script>"
	conn.execute "update �Τ� set �O�@=false where id="&info(9)
	conn.execute "update �Τ� set �O�@=false where id="&id
	diantiao="�P�N��D�G<bgsound src=wav/pk.wav loop=1>["& info(0) &"]�P{"& peiou &"}�F����ĳ�B�}�l��D�A��k�G�ι��Х������Y���I�I�p�w��D�Ӥ���D�A2����������,�x���N������D�I�I<br><marquee width=100% behavior=alternate scrollamount=5><font color=FFD060 size=2>��   "& info(0) &"<a href='tiaozhan1.asp?id=" & info(9)&"' target='d'><img src='../ico/"& info(10) &"-2.gif' width='24' height='24' border='0'></a>["&info(5)&"]---------"& peiou &"<a href='tiaozhan1.asp?id=" & info(9)&"' target='d'><img src='../ico/"& old &"-2.gif' width='24' height='24' border='0'></a>["&menpai&"]    �}�l��Z��D�I��</font></marquee>"
rs.close
rs.open "select top 1 �m�W2,�m�W1 FROM ��D WHERE �m�W2='�L' or �m�W1='�L'",conn
'conn.execute "insert into ��D(�m�W2) values ('"&info(0) &"')"
conn.execute "update ��D set �m�W2='"&info(0)&"'" 
conn.execute "update ��D set �m�W1='"&peiou&"'"
else
	Response.Write "<script language=JavaScript>{alert('�ڤ~���z�O�I');}</script>"
	conn.execute "update �Τ� set �y�O=�y�O-1000 where id="&id
	diantiao="�i�ڵ���D�j"&info(0)&"��"&peiou&"���G��F�A�A�F�`�ڪA�F�A�A�U�����ڪZ�\�m���A�ӻP�A�M�@����  !{"& info(0) &"}�V["& peiou &"]�ڵ���D�A�y�O�U�N1000�I!"
rs.close
rs.open "select top 1 �m�W1,�m�W2 FROM ��D ",conn
conn.execute "update ��D set �m�W1='�L',�m�W2='�L'"
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
sd(194)="����"
sd(195)=info(0)
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="��"
sd(199)="<font color=FFD060>�i��������j</font>"&diantiao
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application("dantiao")=""
Application.UnLock
%>
