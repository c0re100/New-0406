<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ� from �Τ� where id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<20 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=20-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');window.close();}</script>"
	Response.End
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
tx="../ico/"& info(10) &"-2.gif"

action=Request.Querystring("action")
if action="attack" then
Randomize timer
points1=int(100*rnd)+1
points2=int(100*rnd)+1
r=rnd+rnd+1
if points1=points2 then points2=points2+1
Response.Cookies("points1")=points1
Response.Cookies("points2")=points2
Response.Cookies("attack")=""
Response.write "<script language='javascript'>alert('�|������Ƶo�ͩO�H');location.href='gs.asp?"& r &"';</script>"
else
hp1=Request.Cookies("hp1")
hp2=Request.Cookies("hp2")
points1=Request.Cookies("points1")
points2=Request.Cookies("points2")
attack=Request.Cookies("attack")
if hp1="" then hp1=100
if hp2="" then hp2=100
if hp1<=0 then
   response.cookies("hp1")=""
   response.cookies("hp2")=""
   Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
   conn.execute("update �Τ� set �Ȩ�=�Ȩ�+1000000,�ާ@�ɶ�=getdate() where id="&info(9))
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
sd(199)="<font color=red>�i���~�����j["&info(0)&"]</font>�D�ԩ��~�A�g�L�@�f�E�����\���˩��~��o���100�U!(�Q�ȿ��N�ӳo)" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
   response.write "<script language='javascript'>alert('���~�Q���ˤF�A��o�Ȩ�100�U�I�I�I');window.close();</script>"
end if
if hp2<=0 then
   response.cookies("hp1")=""
   response.cookies("hp2")=""
   response.write "<script language='javascript'>alert('�ڳQ���ˤF�I�I�I');window.close();</script>"
end if
%>
<html>
<head>
<title>�D�ԩ��~�a��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px; line-height: 22px}
input {  font-family: "�s�ө���"; font-size: 12px; background-color: #FFFFCC; color: #660000; border: #660000; border-style: dotted; border-top-width: thin; border-right-width: thin; border-bottom-width: thin; border-left-width: thin}
a:link {  font-family: "�s�ө���"; font-size: 12px; color: #660000; text-decoration: none}
a:visited {  font-family: "�s�ө���"; font-size: 12px; color: #660000; text-decoration: none}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#660000" link="#660000" vlink="#660000" background="../ljimage/11.jpg">
<br>
<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td height="25">�D�ԩ��~�a�� �]�ӧQ����100�U��A��²��F�a�I�I�I�^</td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td width="30%" height="37">���~<img src="../IMAGES/010.gif" width="79" height="43"> 
      �ͩR�I 
      <input type="text" readonly size="4" value="<%=hp1%>">
  </td>
    <td width="40%" height="37">&nbsp;</td>
    <td width="30%" height="37">�y����<img src="<%=tx%>" >�ͩR�I 
      <input type="text" readonly size="4" value="<%=hp2%>">
    </td>
  </tr>
  <tr align="center"> 
    <td height="64">&nbsp;</td>
    <td height="64"> <%
Randomize timer
r=rnd+rnd+1
if (points1<>"" or points2<>"") and attack="" then
      if points1<points2 then
         response.cookies("hp1")=hp1-10
         response.cookies("attack")="yes"
         response.write "�Ǫ����G<font color=red>"&points1&"</font> VS �y���̫��G<font color=red>"&points2&"</font><br><a href='gs.asp?"& r &"'>�I����s�ͩR�I</a>"
      else
         response.cookies("hp2")=hp2-10
         response.cookies("attack")="yes"
         response.write "�Ǫ����G<font color=red>"&points1&"</font> VS �y���̫��G<font color=red>"&points2&"</font><br><a href='gs.asp?"& r &"'>�I����s�ͩR�I</a>"
      end if
end if
%> </td>
    <td height="64">
      <input type="button" name="Button" value="�I�Ƨ���" OnClick="javasript:location.href='gs.asp?action=attack&<%=r%>';">
      <input type="button" name="Button" value="�h�X�Գ�" OnClick="javascript:window.close();">
    </td>
  </tr>
</table>
<table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
  <tr align="center"> 
    <td height="40">����������ϥ��s������ Cookie �\��A�Ф��n�������\��<br>
    </td>
  </tr>
</table>
</body>
</html>
<%end if%>