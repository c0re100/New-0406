<%'�D�B
function qiuhun(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�D�B�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if

if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n2�ťH�W�~�i�H�D�B�I');}</script>"
	Response.End
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('���F���@���򪺤����A�x���H�����i�H���B�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �t��,�D�w,�y�O,�Ȩ� FROM �Τ� WHERE id="&info(9),conn
if rs("�t��")<>"�L" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('����O�@�Ҥ@�d��A�A�Q�@����I');}</script>"
	Response.End
end if
if rs("�D�w")<300 or rs("�y�O")<300 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�D�w�P�y�O����300�A�H�a�ݤ��W�A���I');}</script>"
	Response.End
end if
if rs("�Ȩ�")<50000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�S��5�U�����A����n�O���B���I');}</script>"
	Response.End
end if
rs.close
rs.open "select �ʧO,�t��,����,���� FROM �Τ� WHERE �m�W='"&toname&"'" ,conn
if rs("����")="�x��" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���F���@����޲z�B�����v��,�x���T��B�I');}</script>"
	Response.End
end if
if rs("�ʧO")=info(1) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���F�a�A���򤣳\�P���ʪ��I');}</script>"
	Response.End
end if
if rs("�t��")<>"�L" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�Q�@����A�ĤT�̴����r�I');}</script>"
	Response.End
end if
if rs("����")<=1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���@�Ӧ���1�Ū��H�A�S���l�a�I');}</script>"
	Response.End
end if

conn.execute "update �Τ� set �Ȩ�=�Ȩ�-50000 where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application("ljjh_hunyin")=toname
if info(1)="�k" then
qiuhun="["&info(0)&"]�V{"&toname&"}�D�B�G<img src='img/29.gif'>"&fn1&"  <input type=button value='���@�N' onClick=javascript:;disabled=1;window.open('jiehun.asp?id="&info(9)&"&yn=1','d')><input type=button value='���@�N' onClick=javascript:;disabled=1;window.open('jiehun.asp?id="&info(9)&"&yn=0','d')>"
else

qiuhun="["&info(0)&"]�V{"&toname&"}�D�B�G<img src='img/girl.gif'>"&fn1&"  <input type=button value='���@�N' onClick=javascript:;window.open('jiehun.asp?id="&info(9)&"&yn=1','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'><input type=button value='���@�N' onClick=javascript:;window.open('jiehun.asp?id="&info(9)&"&yn=0','d');this.disabled=1 style=background-color:#8AAFF1;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF'>"
end if
end function
%>
