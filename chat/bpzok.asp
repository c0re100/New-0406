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
too=trim(Application("ljjh_bpz"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���� FROM �Τ� WHERE id="&info(9),conn
menpai=rs("����")
rs.close
rs.open "select ����,�m�W FROM �Τ� WHERE id="&id,conn
menpai1=rs("����")
name=rs("�m�W")
if info(0)<>trim(Application("ljjh_bpz")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�O���I�I');}</script>"
	Response.End
end if
rs.close
rs.open "select * FROM �����j�� WHERE �Ĺ�����='"&menpai&"'",conn
if not rs.bof or not rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g�ѥ[�����j�ԤF�I');}</script>"
	Response.End

end if
if yn=1 then
sql="insert into �����j�� (�D������,�Ĺ�����) values ('"&menpai1&"','"&menpai&"')"
conn.Execute(sql)
	Response.Write "<script language=JavaScript>{alert('�}�ԴN�}�ԡI');}</script>"
	jiemen="<font color=B0D0E0>["& menpai &"]</font>�P<font color=B0D0E0>["& menpai1 &"]</font>�F����ĳ�A�w�����20�G00�}�l�����j�ԡA����������̤l���n�ǳơI�I�I"
else
	Response.Write "<script language=JavaScript>{alert('�n�n���A�O�x�F�I');}</script>"
	Response.End

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
sd(195)="�j�a"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i�����P���j</font>"&jiemen
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Application("ljjh_bpz")=""
%>
 