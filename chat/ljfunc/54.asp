<%
function grdb(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('���H��չﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H���H��աI');}</script>"
	Response.End
end if
if fn1<>1 and fn1<>2 and fn1<>3 then
Response.Write "<script language=JavaScript>{alert('��J���~,���ӬO1-3���Ʀr�I');}</script>"
Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �t��,�D�w,�y�O,�Ȩ� FROM �Τ� WHERE id="&info(9),conn
if rs("�D�w")<300 or rs("�y�O")<300 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�D�w�P�y�O����300�A�H�a���M�A��I');}</script>"
Response.End
end if
if rs("�Ȩ�")<1000000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�S��100�U�����A�����աI');}</script>"
Response.End
end if
rs.close
rs.open "select �Ȩ� FROM �Τ� WHERE �m�W='"&toname&"'" ,conn
if rs("�Ȩ�")<1000000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�L�S��100�U�����I');}</script>"
Response.End
end if
rs.close
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000000 where id="&info(9)
set rs=nothing
conn.close
set conn=nothing
Application("ljjh_kissly")=toname
Application("ljjh_bingwen")=fn1
grdb="["&info(0)&"]�V<bgsound src=wav/Faqian.wav loop=1>{"&toname&"}���X��խn�D�A��`100�U�ժ�᪺�Ȥl�I<input type=button value='���Y' onClick=window.open('dubo.asp?id="&info(9)&"&db=1','d')><input type=button value='�Ťl' onClick=window.open('dubo.asp?id="&info(9)&"&db=2','d')><input type=button value='��' onClick=window.open('dubo.asp?id="&info(9)&"&db=3','d')>"
end function
%>