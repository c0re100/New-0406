<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="config.asp"-->
<!--#include file="inc/char.asp"-->
<!--#include file="live.txt"-->

<%
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'JR Wish Board Bottle
'(�馡���@�O�Τ֤k���@��) �� ���@�}�y�~
'�{�ǧ@�� SLIGHTBOY
'���v�Ҧ� SLIGTHBOY
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�Фŭק糧�{��!
'�p�G�T���S���]�A�ݭn�ק� ���pô SLIGHTBOY webmaster@slightboy.com
'�@���o�{���H�Ϩ�ĳ�̡A�������|���ѥ���޳N����A�åB����o�{�Ǫ����򪩥��M�����X�~����L�{�ǡC
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�n���ܶq
Dim PROName,PROEType,PROCNAME,VerNum,VerUpdate,PROVer,PROType,PROMaster,PROPublish,PROCopyRight,CopyRightInfo,JavaCopyRightInfo


PROName="JR Wish Board Bottle"
PROEType="Single-User Personnal"
PROCNAME="(�馡���@�O�Τ֤k���@��) �� ���@�}�y�~"
VerNum="2.0"
VerUpdate="020820"
PROVer=VerNum&" U "&VerUpdate
PROType="�ӤH�Τ᥿���� (²�餤��)"
PROMaster="SLIGHTBOY"
PROPublish=" HTTP://WWW.SLIGHTBOY.COM"
PROCopyRight="CopyRight (c) 2000-2002 SLIGHTBOY.com.All Rights Reserved"
PROCopyRightInfo="�дL���ۧ@�̳Ұ� �O�d�H�W���v�H�� ���¦X�@"

CopyRightInfo = PROName&chr(13)&_
               "����W�r�G"&PROCNAME&chr(13)&_
               "�������G"&PROVer&chr(13)&_
               "���������G"&PROType&chr(13)&_
               "�{�Ǩ�@�G"&PROMaster&chr(13)&_
               "�{�ǰߤ@�����o���a�}:"&PROPublish&chr(13)&_
               PROCopyRight&chr(13)&_
               PROCopyRightInfo

Select Case Request("action")
  Case "click"
   Call SLIGHTBOY_Click
  Case "add"
   Call SLIGHTBOY_Add
  Case "post"
   Call SLIGHTBOY_Post
  Case "super"
   Call SLIGHTBOY_Super
  Case "admin"
   Call SLIGHTBOY_Admin
  Case "del"
   Call SLIGHTBOY_Del
  Case Else
   Call SLIGHTBOY_Look
End Select

Sub SLIGHTBOY_Look
Call HtmlStart
%>
<td align="right" valign="top"><select name="SLIGHTBOY" style="font-size:7pt;font-family:arial" onChange="location.href=this.options[this.selectedIndex].value;" STYLE="BACKGROUND:<%=tdcolor%>;COLOR:<%=trcolor%>">
<%
'�إ߼ƾڮw�s��
Set Conn=server.CreateObject("adodb.connection")
Conn.Open "provider=microsoft.jet.oledb.4.0; data source="&DBpath
'�إ߬��ʹC��
Sql = "Select id,name,sex,age,birth,live,counter From wish Order BY id DESC"
'             0   1    2   3      4    5
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open Sql,conn,1,1

TotalRecord=Rs.RecordCount
Rs.PageSize=Show
MaxPage=Rs.pagecount

OrderPage=Clng(Request("page"))

IF OrderPage="" or OrderPage=0 then
   OrderPage=1
ElseIF OrderPage > MaxPage Then
   OrderPage=MaxPage
End IF

Response.Write "<option value='wish.asp?page="&OrderPage&"' selected Style=""background:"&trcolor&";color:"&tdcolor&""">PAGE"&OrderPage&"</option>"
 For i= 1 TO MaxPage
  Response.Write "<option value='wish.asp?page="&i&"'>PAGE"&i&"</option>"
 Next
%>
</select>�@ <font color="<%=tdcolor%>"><b><%=TotalRecord%></b></font> �� | �y�k: <font color="<%=tdcolor%>"><b><%=use_html%></b></font> | �C�� <font color="<%=tdcolor%>"><b><%=show%></b></font> ��
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="400" background="pic/bg.jpg" height="300" style="border:1 solid black">
<tr>
<td height="230">
<%
IF Session("BottleAdmin")=True Then
Response.Write "<form method=""post"" action=""wish.asp?action=del"">"
End IF

IF not Rs.eof Then
IF OrderPage <> 1 Then
Rs.Move (OrderPage-1)*Show
End IF
i=1

randomize

Do While not Rs.eof and i<=Show

randnum = Int(8*rnd+1)

IF Rs("sex")="m" Then
sexpic="��"
Else
sexpic="��"
End IF
IF Session("BottleAdmin")=True Then
Response.Write "<input type=""checkbox"" name=""Del"" value="""&Rs("id")&""">"
End IF
Response.Write "<marquee scrollamount="""&randnum&""" delay=""0"">"&_
"<a href=""wish.asp?action=click&number="&Rs("id")&""" ONMOUSEOVER=""pop('���~"&Rs("age")&"����b"&Rs("live")&"��"&Rs("name")&sexpic&"�\�F�@���@...<br>��"&Rs("counter")&"�H�ݹL','"&trcolor&"')"";  ONMOUSEOUT=""kill()"">"&_
"<img src=""pic/c"&Rs("birth")&".gif"" border=""0""></a>"&_
"</marquee>"&chr(13)
i=i+1
Rs.MoveNext
Loop
End IF
Rs.Close
Set Rs=Nothing
conn.Close
Set conn=Nothing
%>
<%
IF Session("BottleAdmin")=True Then
Response.Write "<p align=center>"
Response.Write "<input type=""submit"" value="" �R�� "" style=""HEIGHT:22PX;BACKGROUND-COLOR:"&trcolor&";BORDER:1 SOLID BLACK"">"&chr(13)
Response.Write "</form>"
End IF
Call HtmlEnd
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SLIGHTBOY_Click
Call HtmlStart
%>
<td align="right" valign="top">
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="400" background="pic/bg.jpg" height="300" style="border:1 solid black">
<tr>
<td height=230>
<%
'�إ߼ƾڮw�s��
Set Conn=server.CreateObject("adodb.connection")
Conn.Open "provider=microsoft.jet.oledb.4.0; data source="&DBpath
'�إ߬��ʹC��
Sql = "Select * From wish Where id="&Clng(Request.Querystring("number"))
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open Sql,conn,1,3
birth=Rs("birth")
msg=Rs("msg")
IF Rs("sex")="m" Then
sexpic="����"
Else
sexpic="�p�j"
End IF
IF Rs("email")<>"" Then
info= Rs("age")&"���� <a href=""mailto:"&Rs("email")&""">"&Rs("name")&"</a> "&sexpic&",�Ӧ�"&Rs("live")&",��"&Rs("date")&" �d�U���@"
Else
info= Rs("age")&"���� "&Rs("name")&" "&sexpic&",�Ӧ�"&Rs("live")&",��"&Rs("date")&" �d�U���@"
End IF
Rs("counter")=Rs("counter")+1
Rs.Update
Rs.Close
Set Rs=Nothing
conn.Close
Set conn=Nothing
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height=180>
<tr>
<td width="400" style="background-color:white;Filter:Alpha(opacity=40)">
<table border=0 align=center>
<tr>
<td rowspan=2 valign=top>
<img src="pic/c<%=birth%>.gif" align=absmiddle></td>
<td><%=info%></td>
</tr>
<tr>
<td><%=msg%></td>
</tr>
</table>
</td>
</tr>
</table>
<%
Call HtmlEnd
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SLIGHTBOY_Admin
'IF Request.Form("id")="" Then Error("�z�ѤF��g�W�r")
'IF Request.Form("passid")="" Then Error("�z�ѤF��g�K�X")
IF Request.Form("id")<>userid or Request.Form("passid")<>userpass Then
Error("�K�X���~")
Else
Session("BottleAdmin")=True
Response.Redirect "wish.asp"
End IF
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SLIGHTBOY_Post
IF Request.Form("name")="" Then Error("�z�ѤF��g�W�r")
IF Instr("'",Request.Form("name")) or Instr("""",Request.Form("name")) Then Error("�z�W�r�a���D�k�r��<br>PS�G�W�r������a�� ' �Ϊ� """)
IF Request.Form("sex")="" Then Error("�z�ѤF��ܩʧO")
IF CkAge(Request.Form("age")) = False Then Error("�~���氦���J�Ʀr")
IF Request.Form("live")="" Then Error("�z�ѤF��ܩ~��a")
IF Request.Form("birth")="" Then Error("�z�ѤF����@�����O")
IF Request.Form("msg")="" Then Error("�z�ѤF��g�@��")
IF maxmessage = "" or maxmessage > 255 Then maxmessage=255
IF Len(Request.Form("msg"))> maxmessage Then Error("�z���@��Ӫ��F")

IF use_html="ON" Then
msg=HTMLcode(Request.Form("msg"))
Else
msg=HTMLEncode(Request.Form("msg"))
End IF

IF usr_fltbadword="ON" Then msg=CkBadWords(msg)

StrEmail=Request.Form("email")
IF usr_emailcheck="ON" and StrEmail<>"" Then
   IF IsValidEmail(StrEmail)=False Then Error("�п�J���T�H�c�a�}")
End IF

Set Conn=server.CreateObject("adodb.connection")
Conn.Open "provider=microsoft.jet.oledb.4.0; data source="&DBpath
Sql = "Select TOP 1 * From wish Order BY id DESC"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open Sql,conn,3,3
IF Rs.Eof Then
Rs.AddNew
Rs("name")=Trim(Replace(Request.Form("name"),"'",""))
Rs("sex")=Request.Form("sex")
Rs("age")=Request.Form("age")
Rs("birth")=Request.Form("birth")
Rs("email")=StrEmail
Rs("live")=Request.Form("live")
Rs("msg")=msg
Rs("date")=now+time_ctrl
Rs.Update
Else
 IF Rs("msg")=msg Then
 Error("�Фŭ��еn��")
 Else
 Rs.AddNew
 Rs("name")=Request.Form("name")
 Rs("sex")=Request.Form("sex")
 Rs("age")=Request.Form("age")
 Rs("birth")=Request.Form("birth")
 Rs("email")=StrEmail
 Rs("live")=Request.Form("live")
 Rs("msg")=msg
 Rs("date")=now+time_ctrl
 Rs.Update
 End IF 
End IF
Rs.Close
Set Rs=Nothing
conn=Close
Set conn=Nothing

Call HtmlStart
Response.Write "<td align=""right"" valign=""top"">"&chr(13)&_
"</td>"&chr(13)&_
"</tr>"&chr(13)&_
"</table>"&chr(13)&_
"</center>"&chr(13)&_
"</div>"&chr(13)&_
"<div align=""center"">"&chr(13)&_
"<center>"&chr(13)&_
"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""400"" background=""pic/bg.jpg"" height=""300"" style=""border:1 solid black"">"&chr(13)&_
"<tr>"&chr(13)&_
"<td height=230>"&chr(13)&_
"<p align=center><a href=""wish.asp""><img src=""pic/wishing.gif"" width=""95"" height=""90"" border=""0""></a></p>"
Call HtmlEnd
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SLIGHTBOY_Del
IF Session("BottleAdmin")<>True Then Error("���n�o�ö�")

IF Request("del")<>"" Then
DeleteID=Split(Request("del"),",")
Set Conn=server.CreateObject("adodb.connection")
Conn.Open "provider=microsoft.jet.oledb.4.0; data source="&DBpath
For i = 0 to Ubound(DeleteID)
  Sql="delete from wish where id="&DeleteID(i)
  conn.execute(Sql)
Next
conn.close
Set conn=Nothing
End IF
Response.Redirect"wish.asp"
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SLIGHTBOY_Add
Call HtmlStart
%>
<td align="right" valign="top">
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="400" background="pic/bg.jpg" height="300" style="border:1 solid black">
<tr>
<td height=230>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="200">
<tr>
<td align="center" width="400" style="background-color:white;Filter:Alpha(opacity=40)">
<form method="post" action="wish.asp?action=post">
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="1">
<tr>
<td align="right">�W�r :</td>
<td><input type="text" size="12" maxlength="12" name="name" style="COLOR: blue"></td>
<td align="right">�ʧO :</td>
<td>
<select name="sex" size="1">
 <option selected value="">-�п��-</option>
 <option value="m">�k  ��</option>
 <option value="w">�k  ��</option>
</select>
</td>
</tr>
<tr>
<td align="right">�~�� :</td>
<td><input type="text" size="12" name="age" style="COLOR: red"></td>
<td align="right">���O :</td>
<td>
<select name="birth" size="1">
 <option selected value="">-�п��-</option>
 <option value="love">��  �R</option>
 <option value="study">��  �~</option>
 <option value="health">��  �d</option>
 <option value="family">�a  �x</option>
 <option value="work">��  �~</option>
 <option value="future">�N  ��</option>
 <option value="wealth">�]  �I</option>
 <option value="life">��  ��</option>
</select>
</td>
</tr>
<tr>
<td align="right">�H�c :</td>
<td><input type="text" size="12" name="email" style="COLOR: blue"></td>
<td align="right">�a�� :</td>
<td>
<SELECT NAME="live">
<option value="" selected>-�п��-</option>
<%
Location=Split(Location,"|")
For i = 0 to Ubound(Location)
Response.Write "<option value="&Location(i)&">"&Location(i)&"</option>"&chr(13)
Next
%>
</select>
</td>
</tr>
<td colspan="4"><textarea name="msg" rows="4" cols="30"></textarea><font face="�s�ө���"><input type="submit" value="���@" style="BORDER: 1 solid black; COLOR: ffffff; BACKGROUND-COLOR: 808080"></font></td>
</tr>
</table>
</center>
</div>
</form>
</td>
</tr>
</table>
<%
Call HtmlEnd
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SLIGHTBOY_Super%>
<html>
<head>
<title>�u�W�޲z</title>
<%Call HtmlStart%>
<td align="right" valign="top">
</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="400" background="pic/bg.jpg" height="300" style="border:1 solid black">
<tr>
<td height=230>
<table border="0" cellpadding="0" cellspacing="0" width="100%" height="180">
<tr>
<td align="center" width="400" style="Filter:alpha(opacity=40);background-color:white">
<form action="wish.asp?action=admin" method="POST">
<div align="center">
<center>
<table border="0" cellspacing="0">
<tr>
<td align="center" colspan="3">�޲z�M��    �Y�D�޲z�������}����</td>
</tr>
<tr>
<td>�޲z���W��</td>
<td><input type="text" size="15" name="id"></td>
<td>&nbsp;</td>
</tr>
<tr>
<td>�޲z�αK�X</td>
<td><input type="password" size="15" name="passid"></td>
<td><input type="submit" value="�T�w" style="HEIGHT:22PX;BACKGROUND-COLOR:<%=trcolor%>;BORDER:1 SOLID BLACK"></td>
</tr>
</table>
</center>
</div>
</form>
</td>
</tr>
</table>
<%
HtmlEnd
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HtmlStart%>
<html>
<head>
<title><%=title%></title>
<%Call HtmlStyle%>
<META http-equiv=Content-Type content="text/html; charset=big5">
</head>
<body bgcolor="<%=trcolor1%>" text="<%=trcolor2%>" background="<%=bggif%>">
<DIV ID="topdeck" CLASS="popper">&nbsp;</DIV>
<SCRIPT>

var nav = (document.layers); 
var iex = (document.all);
var skn = (nav) ? document.topdeck : topdeck.style;
if (nav) document.captureEvents(Event.MOUSEMOVE);
document.onmousemove = get_mouse;

function pop(msg,bak) 
{

var content ="<TABLE WIDTH=130 BORDER=0 CELLPADDING=2 CELLSPACING=2 BGCOLOR=#808080><TR><TD><FONT COLOR=#FFFFFF>"+msg+"</CENTER></FONT></TD></TR></TABLE>";

  if (nav) 
  { 
    skn.document.write(content); 
	  skn.document.close();
	  skn.visibility = "visible";
  }
    else if (iex) 
  {
	  document.all("topdeck").innerHTML = content;
	  skn.visibility = "visible";  
  }
}

function get_mouse(e) 
{
	var x = (nav) ? e.pageX : event.x+document.body.scrollLeft; 
	var y = (nav) ? e.pageY : event.y+document.body.scrollTop;
	skn.left = x - 60;
  skn.top  = y+20;
}

function kill() 
{
  skn.visibility = "hidden";
}

</SCRIPT>
<p align="center"><a href="<%=userhomepage%>" title="��^ ���� - <%=title%>"><img src="<%=titlegif%>" border="0"></a></p>
<div align="center">
<center>
<table border="0" cellpadding="0" cellspacing="0" width="400">
<tr>
<td><a href="wish.asp"><img src="pic/b1.gif" width="50" height="20" border=0></a><a href="javascript:history.back()"><img src="pic/b2.gif" width="50" height="20" border=0></a><a href="wish.asp?action=super"><img src="pic/b3.gif" width="50" height="20" border=0></a></td>
<%End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HtmlEnd
%>
</td>
</tr>
<tr>
<td valign="bottom" height=70><a href="wish.asp?action=add"><img src="pic/wish.gif" alt="���@" width="72" height="64" border=0></a></td>
</tr>
</table>
</td>
</tr>
</table>
</center>
</div>
<%
Response.Write "<!-- �ФŧR�� ���v�n�� -->"&chr(13)&_
"<div align=""center""><center>"&chr(13)&_
"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""400"">"&chr(13)&_
"<tr>"&chr(13)&_
"<td align=""right"">"&chr(13)&_
"<img src=""pic/cr.gif"" width=""280"" height=""9"" border=""0"" title="""&CopyRightInfo&"""> <a href=""http://www.slightboy.com""><img src=""pic/slightboy.gif"" alt="""&CopyRightInfo&""" width=""12"" height=""12"" border=""0""></a>"&chr(13)&_
"</td>"&chr(13)&_
"</tr>"&chr(13)&_
"</table>"&chr(13)&_
"</center>"&chr(13)&_
"</div>"&chr(13)&_
"<!-- �ФŧR�� ���v�n�� -->"&chr(13)
%>
</form>
</body>
</html>
<%
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub HtmlStyle%>
<STYLE TYPE="text/css">
<!--
body,td,input,select   {font-family:�s�ө���;font-size:9pt;}
textarea,input,button  {font-family:�s�ө���;font-size:9pt;border:1 solid <%=tdcolor%>;background-color:<%=trcolor%>;color:<%=tdcolor%>}
a:hover { font-family:�s�ө���;text-decoration:underline; color:<%=trcolor%>;}
a { font-family:�s�ө���;text-decoration:none; color:<%=tdcolor%>;}

body{overflow:scroll;overflow-x:hidden}
.popper{position : absolute;visibility : hidden;}
//-->
</style>
<%End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub Error(msg)
Call HtmlStart
Response.Write "<td align=""right"" valign=""top"">"&chr(13)&_
"</td>"&chr(13)&_
"</tr>"&chr(13)&_
"</table>"&chr(13)&_
"</center>"&chr(13)&_
"</div>"&chr(13)&_
"<div align=""center"">"&chr(13)&_
"<center>"&chr(13)&_
"<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""400"" background=""pic/bg.jpg"" height=""300"" style=""border:1 solid black"">"&chr(13)&_
"<tr>"&chr(13)&_
"<td height=230>"&chr(13)&_
"<p align=center><br><br><font color=D00000>"&msg&"</font></p>"
Call HtmlEnd
Response.End
End Sub
%>