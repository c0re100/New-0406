<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
my=info(0)
if my="" then response.redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs=conn.execute("select ��O from �Τ� where id="&info(9))
tl=rs("��O")
if tl<0 then
  conn.execute("update �Τ� set ���A='��' where id="&info(9))
  Session.Abandon
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
  response.write "<script language='javascript'> alert('�A�n�����F�ڡI�I�I'); window.close(); </script>"
    response.end
end if
set rs=nothing
points=session("points")
rpg=session("RPG")
if points="" then points=0
a=Request.Querystring("a")
if a="tou" then
  rpg_times=Request.cookies("rpg_times")
  if rpg_times="" then rpg_times=0
  if rpg_times>3 then

	set rs=nothing	
	conn.close
	set conn=nothing  
    response.write "<script language='javascript'> alert('�@�Ѱ��઱�T���I�I�I'); window.close(); </script>"
   response.end
  end if
  rpg_times=rpg_times+1
  Randomize timer
  s=rnd/(rnd+1)+1
  r=int(rnd*8)+1
  points=points+r
  if points>70 then
     points=0
     response.cookies("rpg_times")=rpg_times
     response.cookies("rpg_times").expires=now()+1
  end if
  session("points")=points
  session("RPG")=""
  response.write "<script language='javascript'> alert ('�A��F"& r &"�I'); location.href='index.asp?"& s &"'; </script>"
  response.end
else
  Randomize timer
  r=rnd/(rnd+1)+1
%>
<html>
<head>
<title>�F�C����--�ϪѤl</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "verdana"; font-size: 12px}
a {  cursor: help}
-->
</style>
</head>
<body bgcolor="#fffddf" text="#660000">
<table width="100%" border="1" cellspacing="5" cellpadding="3" bordercolordark="#333333" bordercolorlight="#FFFFFF">
  <tr align="center"> 
    <td colspan="3" height="30">�b�~���X�{<img src="h.gif" width="32" height="32">�N���ƥ�o�� 
      �ثe�a�I�G<font color=red><%=points%> </font>&nbsp;&nbsp;�ϪѤl<a href="index.asp?a=tou&<%=r%>"><img src="run.gif" width="38" height="36" border="0"></a><br>
      <br>
      �`�N�G�����_�C���ள��@��I�I�I</td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">0.</font>�X�o�I�G���n�B�I</td>
    <td width="33%"><font color="#000066">25.</font>����G <%
if points=25 then
 if rpg="" then
      conn.execute("update �Τ� set �Ȩ�=�Ȩ�+1000 where id="&info(9))
      session("RPG")=true
%> <img src="h.gif" width="32" height="32">�o�{1000��Ȥl�I�I�I <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">50.</font>����G<%
if points=50 then
 if rpg="" then
  set rs=conn.execute("select id from ���~ where ���~�W='�P���C' and �֦���="&info(9))
    if rs.eof and rs.bof then
     conn.execute("insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,sm,�ƶq,�Ȩ�,����,�|��) values('�P���C',"&info(9)&",'����',0,0,0,5000,10000,2004114,2004114,1,8000000,120,"&aaao&")")
else
id=rs("id")
 conn.execute("update ���~ set �ƶq=�ƶq+1 where id="&id)
end if
 rs.close : set rs=nothing
    session("RPG")=true
%><img src="h.gif" width="32" height="32">�o�{�P���C�I�I�I <%
end if
end if%> </td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">5.</font>����G <%
if points=5 then
if rpg="" then
conn.execute("update �Τ� set �Ȩ�=�Ȩ�+5000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">�o�{5000��Ȥl�I�I�I <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">28-29.</font>�Z�]�G <%
if points=28 or points=29 then
if rpg="" then
conn.execute("update �Τ� set ���O=���O+500 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">���v�ǪZ���O����300�I�I�I <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">54-55.</font>�a�V�G <%
if points=54 or points=55 then
if rpg="" then
conn.execute("update �Τ� set ��O=��O-1000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">�^���a�V����O1000�I�I�I <%
end if
end if%> </td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">9.</font>���@�T���G <%
if points=9 then
if rpg="" then
conn.execute("update �Τ� set �Ȩ�=�Ȩ�-5000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">�ᥢ5000��Ȥl�I�I�I <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">35.</font>�K�|�G <%
if points=35 then
if rpg="" then
conn.execute("update �Τ� set �k�O=�k�O+2 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">�֬��F�@�f�A�k�O����2�I�I�I <%
end if
end if%></td>
    <td width="34%"><font color="#000066">60.</font>�a�V�G <%
if points=60 then
if rpg="" then
set rs=conn.execute("select id from ���~ where ���~�W='�۫�C' and �֦���="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into ���~(���~�W,�֦���,����,���s,��O,���O,��T��,����,����,sm,�ƶq,�Ȩ�,����,�|��) values('�۫�C',"&info(9)&",'�k��',0,0,0,40000,15000,2004113,2004113,1,5000000,140,"&aaao&")")
else
id=rs("id")
 conn.execute("update ���~ set �ƶq=�ƶq+1 where id="&id)
end if
rs.close : set rs=nothing
session("RPG")=true
%><img src="h.gif" width="32" height="32">�^���a�V����O<br>
      500�A���Ƶo�{�F�@���_�C�I�I�I <%
end if
end if%></td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">13-14.</font>�׫B��G <%
if points=13 or points=14 then
if rpg="" then
set rs=conn.execute("select id from ���~ where ���~�W='�p�^���y' and �֦���="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into ���~(���~�W,�֦���,����,���s,��O,���O,��T��,����,����,sm,�ƶq,�Ȩ�,����,�|��) values('�p�^���y',"&info(9)&",'���~',0,0,0,40000,15000,99006,99006,1,10000000,0,"&aaao&")")
else
id=rs("id")
conn.execute("update ���~ set �ƶq=�ƶq+1  where id="& id &" and �֦���="&info(9))
end if
session("RPG")=true
%><img src="h.gif" width="32" height="32">�@�}�{�q�R�E��I�b�@�ʥj�Ѫ��j��W�o�{�@���m�p�^���y�n! <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">38-39.</font>�q���G <%
if points=38 or points=39 then
if rpg="" then
set rs=conn.execute("select id from ���~ where ���~�W='�¥�' and �֦���="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into ���~(���~�W,�֦���,����,����,���s,��O,���O,��T��,����,sm,�ƶq,�Ȩ�,����,�|��) values('�¥�',"&info(9)&",'���~',0,0,0,0,0,2003305,2003305,1,10000,0,"&aaao&")")
else
id=rs("id")
 conn.execute("update ���~ set �ƶq=�ƶq+1 where id="&id)
end if
session("RPG")=true
%><img src="h.gif" width="32" height="32">�b�q���~�}�ߨ�@���¥ۡI�I�I <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">64-65.</font>�]�`�|�G <%
if points=64 or points=65 then
if rpg="" then
conn.execute("update �Τ� set ���O=���O-500 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">���������l�����O500�I�I�I <%
end if
end if%> </td>
  </tr>
  <tr> 
    <td width="33%"><font color="#000066">19-20.</font>�ؿv�u�a�G <%
if points=19 or points=20 then
if rpg="" then
set rs=conn.execute("select id from ���~ where ���~�W='���ݪO' and �֦���="&info(9))
if rs.eof and rs.bof then
  conn.execute("insert into ���~(���~�W,�֦���,����,����,���s,��O,���O,��T��,����,sm,�ƶq,�Ȩ�,����,�|��) values('���ݪO',"&info(9)&",'���~',0,0,0,0,0,2003304,2003304,1,10000,0,"&aaao&")")
else
id=rs("id")
 conn.execute("update ���~ set �ƶq=�ƶq+1 where id="&id)
end if
session("RPG")=true
%><img src="h.gif" width="32" height="32">���[�ؿv�u�@����s���@���u����ݪO�I�I�I <%
end if
end if%> </td>
    <td width="33%"><font color="#000066">45.</font>���ַ|�G <%
if points=45 then
if rpg="" then
conn.execute("update �Τ� set ��O=��O+1000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">��O��_1000�I�I <%
end if
end if%> </td>
    <td width="34%"><font color="#000066">70.</font>�L�H�q�G <%
if points=70 then
if rpg="" then
conn.execute("update �Τ� set �Ȩ�=�Ȩ�+8000 where id="&info(9))
session("RPG")=true
%><img src="h.gif" width="32" height="32">�o�{8000��Ȥl�I�I�I <%
end if
end if%> </td>
  </tr>
</table>
</body>
</html>
<%
end if
conn.close
set conn=nothing
%>