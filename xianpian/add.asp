<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from �Τ� where id="&info(9),conn
if rs.bof or rs.eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�藍�_�A�S���n��!');location.href = '../exit.asp';}</script>"
	Response.End
mycorp=rs("����")
end if%>
<%if request("action")<>"add" then%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta http-equiv="Content-Language" content="zh-cn">
<style type="text/css">
<!--
table {  font-size: 9pt}
td {  font-size: 9pt}
a:link {color:#000000;text-decoration: none}
a:visited {color:#000000;text-decoration: none}
a:active {color:#000000;text-decoration: none}
a:hover {color:red;text-decoration: none}
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�j�L�[�J</title></head>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" topmargin="0" leftmargin="0" background="../ljimage/11.jpg">
<div align="center"><center><form method="POST" action="add.asp"><table border="1" width="600" height="310" cellspacing="0" bgcolor="#DCE2FF" bordercolor="#988CD0">

<tr><td width="594" colspan="2" bgcolor="#FFFFFF">
<p align="center"><font face="�s�ө���" size="3"><b>�Фj�L�ԲӶ�g�A�����</b></font>
</td></tr>
<tr><td width="594" colspan="2">

<div align="center"><center><table border="1" width="100%" bordercolorlight="#000000" cellspacing="0" cellpadding="4" bordercolordark="#FFFFFF" bgcolor="#F8EBD1">
<tr>
<td width="13%" align="center"><a href=main.asp?action=search>��-�j�L�C��</a></td>
<td width="13%" align="center">��j�L�[�J</td>
<td width="13%" align="center"><a href=userre.asp>��-�ק���</a></td>
<td width="13%" align="center"><a href="sf.asp?Keys=Login">��-�W�ź޲z</a></td>
<td width="13%" align="center"><a href="admsearch.asp?keys=admsearch">��-���ŷj��</a></td>
<td width="35%" align="center">
<center>�{��<font color="#FF6600">    <!--#include file="zongshu.asp" -->
</font>��j�L�[�J</center>
</td>
</tr></table></center></div>

</td></tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>�m�W�G</td>
<td width="479" bgcolor="#DCE2FF">
<input name=action type=hidden value=add>
<input type="text" name="name" value="<%=uname%>" size="15" maxlength="12">
</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>�����G</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="mode" value="<%=rs("����")%>" size="15" maxlength="50">�A�b�ڤ۪�����</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>�K�X�G</td>
<td width="479" bgcolor="#DCE2FF"><input type="password" name="password" size="15" maxlength="30">����«ȡA�Ф��n�M�ڤ۪��n���K�X�ۦP</td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>�T�{�K�X�G</td>
<td width="479" bgcolor="#DCE2FF"><input type="password" name="repassword" size="15" maxlength="30">�P�W</td>
</tr>

<tr>
<td width="111" align="right">�ʧO�G</td>
<td width="479" bgcolor="#DCE2FF"><select name="sex" size="1" class="p9">
<option value="�k">�ʧO���k </option>
<option value="�k">�ʧO���k </option>
</select></td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>�H�c�G</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="mail" size="30" maxlength="50" value="<%=rs("�H�c")%>"></td>
</tr>

<tr>
<td width="111" align="right">�D���G</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="url" size="30" maxlength="50" value="http://"></td>
</tr>

<tr>
<td width="111" align="right">�z�������M�I���G</td>
<td width="479" bgcolor="#DCE2FF"><select name="icq_img" class="p9" size="1">
<option value="oicq.gif" selected>OICQ</option>
<option value="icq.gif">ICQ</option>
</select>
<input type="text" name="icqnum" size="10" maxlength="10" value="<%=rs("oicq")%>">
<a href="http://www.icq.com">ICQ�U��(�^��)</a>�A
<a href="http://www.oicq.com">OICQ�U��(����)</a></td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>�z���ͤ�G</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="years" size="4" maxlength="4" value="19" class="p9">�~
<input type="text" name="mons" size="2" maxlength="2">��
<input type="text" name="days" size="2" maxlength="2">��</td>
</tr>

<tr>
<td width="111" align="right">�pô�q�ܡG</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="likes" size="30" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">����ΩI���G</td>
<td width="479" bgcolor="#DCE2FF">
<input type="text" name="gsm" size="30" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">�l�F�s�X�G</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="pc" size="16" maxlength="6"></td>
</tr>

<tr>
<td width="111" align="right">���W�١G</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="danwei" size="50" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">�ԲӦa�}�G</td>
<td width="479" bgcolor="#DCE2FF"><input type="text" name="address" size="50" maxlength="50"></td>
</tr>

<tr>
<td width="111" align="right">�Ӥ��W�ǡG</td>
<td width="479" bgcolor="#ddE2FF"><font color=aa0000>�Х��[�J��A�A��[<a href=userre.asp>�ק���</a>]�B�i��ާ@<vinput type="file" name="photo" class="p9"> </font> </td>
</tr>

<tr>
<td width="111" align="right"><font color="#FF0000">*</font>��p²���G</td>
<td width="479" bgcolor="#DCE2FF"><textarea name="doc" rows="3" cols="39"><%=rs("ñ�W")%></textarea> </td>
</tr>

<tr>
<td width="594" colspan="2">
<center><input type="submit" value="�[�J" name="B1">
&nbsp;&nbsp;&nbsp;
<input type="reset" value="�M��" name="B2"></center></td>
</tr>


</table></center></form></div>

</body>
</html>
<%Else
Dim name,repassword,password,sex,mail,url,icq_img,icqnum,age,gsm,pc,danwei
Dim years,mons,days,likes,address,photo,doc,mode
Dim ip,counter,n,y,r,s,f,m,sj,times,conn,sql,rs,DBPath,strSQL
Dim lists,xingzuo

name     = Server.HtmlEncode(Request.Form("name"))      '�m�W
password = Server.HtmlEncode(Request.Form("password"))  '�K�X
repassword = Server.HtmlEncode(Request.Form("repassword"))  '�T�{�K�X
sex      = Server.HtmlEncode(Request.Form("sex"))       '�ʧO
mail     = Server.HtmlEncode(Request.Form("mail"))      'E-Mail
URL      = Server.HtmlEncode(Request.Form("URL"))       '�����챵
icq_img  = Server.HtmlEncode(Request.Form("icq_img"))   '�ǩI���Ϥ�
icqnum   = Server.HtmlEncode(Request.Form("icqnum"))    '�ǩI�����X
AGE      = Server.HtmlEncode(Request.Form("age"))       '�~��
years    = Server.HtmlEncode(Request.Form("years"))     '�ͤ骺�~
mons     = Server.HtmlEncode(Request.Form("mons"))      '�ͤ骺��
days     = Server.HtmlEncode(Request.Form("days"))      '�ͤ骺��
likes    = Server.HtmlEncode(Request.Form("likes"))     '�q��
gsm      = Server.HtmlEncode(Request.Form("gsm"))       '���
pc  	 = Server.HtmlEncode(Request.Form("pc"))   	'�l�F�s�X
danwei   = Server.HtmlEncode(Request.Form("danwei"))    '���
address  = Server.HtmlEncode(Request.Form("address"))   '�~��a
photo    = Server.HtmlEncode(Request.Form("photo"))     '�ۤ�
doc      = Server.HtmlEncode(Request.Form("doc"))       '�d��
mode     = Server.HtmlEncode(Request.Form("mode"))      '�Z��
ip       = Request.ServerVariables("REMOTE_ADDR")
if repassword<>password  then _
response.write "�K�X���~�C<a href = javascript:history.back()>��^����</a>":response.end

'�ϥ�ServerVariablesŪ���Τ᪺IP�a�}�A�N�Τ�IP��pIP�ܶq��
counter  = "0"
n=Year(date())                                           '���o���Year�~���~����Jn���ܶq��
y=Month(date())                                          '�������Month�몺�����Jy���ܶq��
r=Day(date())                                            '�������Day�Ѫ��ѼƩ�Jr���ܶq��
s=Hour(time())                                           '���o���Hour�p�ɪ��p�ɼƩ�Js�ܶq��
f=Minute(time())                                         '�������Minute���������ƭȩ�Jf�ܶq��
m=Second(time())                                         '�������Second����Ʃ�Jm�ܶq��
if len(y)=1 then y="0" & y                               '�p�Gy���ƭȵ���@��ƪ��ܡA�hy����0+Y�ܶq�A�]�N�O���A���ƫe���[�@��0�A��p01 or 02 or 03
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "�~" & y & "��" & r & "��" & " " & weekdayname(weekday(date())) & " " & s & ":" & f & ":" & m
'sj ���� X�~X��X��Ů� �P���X �X�I�X���X��
'WeekDay�O���Ѥ@��1�P7�������ƭȡA�]�N�O�P���X
times    = sj
'#####################�P�y�P�_#########################
'�s�Ϯy �]12/22-1/19�^ ���~�y �]1/20-2/18�^ �����y �]2/19-3/20�^ �զϮy �]3/21-4/19�^
'�����y �]4/20-5/20�^ ���l�y �]5/21-6/20�^ ���ɮy �]6/21-7/22�^ ��l�y �]7/23-8/22�^
'�B�k�y �]8/23-9/22�^ �Ѭi�y �]9/23-10/22�^ ���Ȯy �]10/23-11/21�^ �g��y �]11/22-12/21�^
if days="" or years="" or mons="" then
response.write "�п�J�ƭ�!<a href = javascript:history.back()>��^����</a>"
response.end
end if
if IsNumeric(mons)=false or IsNumeric(days)=false or IsNumeric(years)=false then
response.write "�z��J�����O�Ʀr�C<a href = javascript:history.back()>��^����</a>"
response.end
end if
mons=cint(mons)
days=cint(days)
years=cint(years)

if mons=>13 or mons<=0 then _
response.write "������~�A�̦h�u��12�Ӥ�A�̤֤]��1��C<a href = javascript:history.back()>��^����</a>":response.end
if days=>32 or days<=0 then _
response.write "������~�A�л{�u�I��g�I�I<a href = javascript:history.back()>��^����</a>":response.end
if years<=1899 or years>=year(now()) then _
response.write "�~���O�q1900�~�}�l�p��A��M�A�A�]�����g2000�~�X�ͪ��C<a href = javascript:history.back()>��^����</a>":response.end

if (mons=12 and days>=22) or (mons=1 and days<=19) then
xingzuo= "�s�Ϯy"
elseif (mons=1 and days>=20) or (mons=2 and days<=18) then
xingzuo= "���~�y"
elseif (mons=2 and days>=19) or (mons=3 and days<=20) then
xingzuo= "�����y"
elseif (mons=3 and days>=21) or (mons=4 and days<=19) then
xingzuo= "�զϮy"
elseif (mons=4 and days>=20) or (mons=5 and days<=20) then
xingzuo= "�����y"
elseif (mons=5 and days>=21) or (mons=6 and days<=20) then
xingzuo= "���l�y"
elseif (mons=6 and days>=21) or (mons=7 and days<=22) then
xingzuo= "���ɮy"
elseif (mons=7 and days>=23) or (mons=8 and days<=22) then
xingzuo= "��l�y"
elseif (mons=8 and days>=23) or (mons=9 and days<=22) then
xingzuo= "�B�k�y"
elseif (mons=9 and days>=23) or (mons=10 and days<=22) then
xingzuo= "�Ѭi�y"
elseif (mons=10 and days>=23) or (mons=11 and days<=21) then
xingzuo= "���Ȯy"
elseif (mons=11 and days>=22) or (mons=12 and days<=21) then
xingzuo= "�g��y"
end if

'response.write "��e�~��:" & year(now())
'response.write "<br>�����p��z�����ƬO�G" & (year(now()) - years)

'###############################
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("cidu-net.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
SQL="select * from list where name='"+name+"' order by ID ASC"
rs.open sql,conn,3,3
if rs.eof and rs.bof  then
elseif name = rs("name") Then
name = ""
Password=""
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�Τ�W�w�s�b�A�Э��s�n����g�I');" & Chr(13)
response.write "</script>" & Chr(13)
response.write "<center><font color=red><a href = javascript:history.back()>��^����</a></font></center>"
Response.End
End If
rs.close



'--------------------�P�_�O�_�g�J����r�q--------
if name="" or mail="" or password="" Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�z�S����g�W�٩αK�X�ΫH�c�A�Э��s�n����g�I');" & Chr(13)
response.write "window.document.location.href=javascript:history.back();"&Chr(13)
response.write "</script>" & Chr(13)
End IF

if icq_img<>"0" Then
If icqnum="" Or IsNumeric(icqnum)=False Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('OICQ��icq���X��J���~!�Э��s��J�I�I�I');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if

end if

if icq_img = "0" then
If icqnum <> "" Then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�п�ܶǩI�������I');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if
if icqnum = "" then
icqnum = "-------"
icq_img= null
end if
End If

if doc="" then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('�п�J�d���I');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
end if

if likes = "" then
likes="--------"
end if

if url = "" then
url="--------"
end if

if address = "" then
address="--------"
end if

if pc = "" then
pc=0
end if

if danwei = "" then
danwei="--------"
end if

if gsm = "" then
gsm="-------"
end if

if photo = "" then
photo="-------"
end if

'--------------------���T��mail�P�_���----------
if IsValidEmail(mail)=False then
response.write "<script language='javascript'>" & chr(13)
response.write "alert('���~��email��J');" & Chr(13)
response.write "window.document.location.href='javascript:history.back()';"&Chr(13)
response.write "</script>" & Chr(13)
End if


function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
IsValidEmail = false
exit function
end if
for each name in names
if Len(name) <= 0 then
IsValidEmail = false
exit function
end if
for i = 1 to Len(name)
c = Lcase(Mid(name, i, 1))
if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
IsValidEmail = false
exit function
end if
next
if Left(name, 1) = "." or Right(name, 1) = "." then
IsValidEmail = false
exit function
end if
next
if InStr(names(1), ".") <= 0 then
IsValidEmail = false
exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
IsValidEmail = false
exit function
end if
if InStr(email, "..") > 0 then
IsValidEmail = false
end if

end function

'-------------�g�J�ƾڮwstart-------------
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("cidu-net.asp")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
'sql="select * from list order by ID ASC"
sql="list"
rs.open sql,conn,3,3
rs.addnew
photo = null
rs.fields("name")=name
rs.fields("password")=password
rs.fields("sex")=sex
rs.fields("mail")=mail
rs.fields("url")=url
rs.fields("icqnum")=icqnum
rs.fields("icq_img")=icq_img
rs.fields("xingzuo")=xingzuo
age=n-years
rs.fields("age")=age       '�~�ֻݭn�ھڥͤ�ק�A�ӥB�ݭn��@�P�y�N�X
rs.fields("years")=years
rs.fields("mons")=mons
rs.fields("days")=days
rs.fields("likes")=likes
rs.fields("gsm")=gsm
rs.fields("pc")=pc
rs.fields("danwei")=danwei
rs.fields("address")=address
rs.fields("photo")=photo
rs.fields("doc")=doc
rs.fields("mode")=mode
rs.fields("ip")=ip
rs.fields("counter")=counter
rs.fields("times")=times
rs.update
rs.close
'-----------------------�ƾڲK�[���\,���������main.asp-----------
Response.Redirect "main.asp?action=search"
End If
%>
