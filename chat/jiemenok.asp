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
too=trim(Application("ljjh_tongmen"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,�P�� FROM �Τ� WHERE id="&info(9),conn
menpai=rs("����")
tongmen=rs("�P��")
rs.close
rs.open "select ����,�P��,�m�W FROM �Τ� WHERE id="&id,conn
menpai1=rs("����")
tongmen1=rs("�P��")
name=rs("�m�W")
if info(0)<>trim(Application("ljjh_tongmen")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�Q�@����A�H�a�]�S�������A�I');}</script>"
	Response.End
end if
if yn=1 then
	Response.Write "<script language=JavaScript>{alert('�J�M�A�o�򦳸۷N�ڦP�N�����F�I');}</script>"
if  Instr(1,tongmen,"|�L|")=0  then
tongmen=tongmen+menpai1&"|"
else
tongmen="|"&menpai1&"|"
end if
if  Instr(1,tongmen1,"|�L|")=0  then
tongmen1=tongmen1+menpai&"|"
else
tongmen1="|"&menpai&"|"
end if
	conn.execute "update �Τ� set �P��='"& tongmen &"' where ����='"&info(5)&"'"
	conn.execute "update �Τ� set �P��='"& tongmen1 &"'  where ����='"&menpai1&"'"

	jiemen="<font color=FFF0D0>["& menpai &"]</font>�P<font color=FFF0D0>["& menpai1 &"]</font>�l�⵲���A�����������ʤ��n�I�I"
   

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
sd(196)="FFF0D0"
sd(197)="FFF0D0"
sd(198)="��"
sd(199)="<font color=FFF0D0>�i�����P���j</font>"&jiemen
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Application("ljjh_tongmen")=""
%>
 