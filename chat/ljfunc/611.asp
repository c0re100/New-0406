<%'�@��
function fakuan(fn1,to1,toname)
if info(2)<8 then
	Response.Write "<script language=JavaScript>{alert('�@�ڻݭn8�ťH�W�޲z���ާ@�I');}</script>"
	Response.End
end if
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�@�ڹﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,�Ȩ� FROM �Τ� WHERE �m�W='"&toname&"'",conn
if info(2)<rs("grade") then
	Response.Write "<script language=JavaScript>{alert('���i�H�ﰪ�ź޲z���ާ@!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("�Ȩ�")<10000 then
	Response.Write "<script language=JavaScript>{alert('���w�g�S���i�H�@�F!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000000 where �m�W='"&toname&"'"
if info(5)<>"�x��" then
fakuan= info(0) &"[�K����]:<font color=red><bgsound src=wav/daipu.wav loop=1>�ѩ�" & toname & "�b���򤤹H�ϳW�h,�x���M�w���@�ڻȨ�100�U!</font>"
else
fakuan= info(0) &"[�x���H��]:<font color=red><bgsound src=wav/daipu.wav loop=1>�ѩ�" & toname & "�b���򤤹H�ϳW�h,�x���M�w���@�ڻȨ�100�U!</font>"
end if 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>
