<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set cn=Server.CreateObject("ADODB.Connection")
Set rst=Server.CreateObject("ADODB.RecordSet")
guess="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("guess.asa")
cn.Open guess

tempdate=date
rst.open"select * from �������X  where  datediff('d','"&tempdate&"',�m�����)=0",cn,1,1
if rst.bof then
RANDOMIZE
guessno=int((99999-10000+1)*rnd+10000)
'guessno=12345
cn.execute"update �������X set �������X='"&guessno&"',�m�����='"&tempdate&"'"
application("guess")=guessno
end if
if application("guess")="" then application("guess")=rst("�������X")
	sql="SELECT * FROM �������X"
	Set Rst=cn.Execute(sql)
	gold=rst("�֭p���B")
%>
<html>
<head>
<title>�F�C�m������</title> <meta http-equiv="Content-Type" content="text/html; charset=big5"> 
<link rel="stylesheet" href="chat/READONLY/Style.css"> <script language="JavaScript">
<!--
function MM_showHideLayers() { //v3.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v='hide')?'hidden':v; }
    obj.visibility=v; }
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script> 
</head>

<body bgcolor="#FF9900" LEFTMARGIN="0" TOPMARGIN="0">
<div id="Layer1" style="position:absolute; left:157px; top:38px; width:278px; height:203px; z-index:1; background-color: #FFFFFF; layer-background-color: #FFFFFF; border: 1px none #000000; visibility: hidden"> 
<table width="100%" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF"> 
<tr> <td> <p align="center"><font size="+1" color="#0000FF">�m������</font></p><p> <font size="2"> ���m���C���12�ɶ}���A�����������ѹq������A�����L�k�ާ@�A���������A�����N����U�@���C�C���@�H���i�ʶR�@�Ӹ��X�A�ʶR���X��5��Ʀr�C�R�@�Ӹ��X��100��Ȥl�A�֭p���N�[1000��C</font></p><p align="right"><font size="2">------�F�C����m������</font></p><p align="center">[<a href="#" onClick="MM_showHideLayers('Layer1','','hide')">���������f</a>]</p></td></tr> 
</table></div><TABLE WIDTH="463" BORDER="0" ALIGN="CENTER"><TR><TD><TABLE BORDER=1 CELLSPACING=0 CELLPADDING=2 ALIGN="center" BORDERCOLORDARK="#FFFFFF" WIDTH="32%" HEIGHT="31"> 
<TR ALIGN="center"> <TD BGCOLOR="#669900" WIDTH="100%" HEIGHT="25"><IMG BORDER="0" SRC="xycpzx.jpg" WIDTH="463" HEIGHT="85"></TD></TR> 
</TABLE><P ALIGN="CENTER"><A HREF="#" ONCLICK="MM_showHideLayers('Layer1','','show')"><FONT COLOR="#FFFFFF">�� 
�R �� ��</FONT></A><BR></P><TABLE WIDTH="263" BORDER="1" CELLSPACING="0" CELLPADDING="3" ALIGN="center" BORDERCOLORLIGHT="#000000" BORDERCOLORDARK="#FFFFFF" BGCOLOR="#FFFFFF"> 
<TR> <TD WIDTH="116">�U�������v�F��</TD><TD WIDTH="129"><%=gold%></TD></TR> <TR> <TD WIDTH="116">�����������X�O</TD><TD WIDTH="129"><%=application("guess")%></TD></TR> 
<TR> <%
sql="SELECT * FROM �ʶR�� where datediff('d','"&tempdate&"',�ʶR���)=-1  and ���X="&application("guess")&""
Set Rst=cn.Execute(sql)
if rst.bof then
ren="�L�H����"
else
ren=rst("�ʶR��")
'�����H�[�Ȩ�
sql="select * from �ʶR�� where �ʶR��='"&ren&"' and ���=false"
set rst=cn.execute(sql)
	if not(rst.bof) then 
	sql="update �ʶR�� set ���=true"
	set rst=cn.execute(sql)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
	sql="update �Τ� set �Ȩ�=�Ȩ�+"&gold&" where �m�W='"&ren&"'"
	set rs=conn.execute(sql)
'�[�J�����W
sql="INSERT into �����H(������,�������B,�������X,�������)values('"&ren&"',"&gold&","&application("guess")&",'"&tempdate&"')"
Set Rst=cn.Execute(sql)
'�֭p���M0
	sql="update �������X set �֭p���B=0"
	set rst=cn.execute(sql)
	end if
end if
	sql="select * from �����H where �������X="&application("guess")&" and ������='"&ren&"'"
	set rst=cn.execute(sql)
	if not(rst.bof) then
	ren=rst("������")
	jin=rst("�������B")
	else
	ren="�L�H����"
	jin="�L�H����"
	end if
	%> <TR> <TD WIDTH="116">���������H�O</TD><TD WIDTH="129"><%=ren%></TD></TR> 
<TR> <TD WIDTH="116">�����������B</TD><TD WIDTH="129"><%=jin%></TD></TR> </TABLE><FORM METHOD="post" ACTION="buyguess.asp"> 
<TABLE WIDTH="235" BORDER="1" CELLSPACING="0" CELLPADDING="3" ALIGN="center" BORDERCOLORLIGHT="#000000" BORDERCOLORDARK="#FFFFFF" BGCOLOR="#FFFFFF"> 
<TR> <TD WIDTH="99">�� �R ��</TD><TD WIDTH="106"><%=info(0)%></TD></TR> <TR> <TD WIDTH="99">�z�����X</TD><TD WIDTH="106"> 
<%
	sql="select * from �ʶR�� where �ʶR��='"&info(0)&"' and datediff('d','"&tempdate&"',�ʶR���)=0"
	set rst=cn.execute(sql)
	if rst.bof then
	%> <INPUT TYPE="text" NAME="guessnumber" SIZE="10" MAXLENGTH="10"> </TD></TR> 
<TR> <TD COLSPAN="2"> <DIV ALIGN="center"> <INPUT TYPE="submit" NAME="Submit" VALUE="�ʶR"> 
   <INPUT TYPE="reset" NAME="Submit2" VALUE="�M��"> </DIV></TD></TR> <%
else
myguess=rst("���X")
response.write myguess
end if
%> </TABLE><FONT COLOR="#FFFFFF" SIZE="2"> <%if ren="�L�H����" then%> �����L�H�����A�����N�֭p��U�@���A�Ʊ�U�@���������̴N�O�z! 
<BR> �����ʶR�A�����ǩ��Ѥ��j�����N�O�A��I <%else%> ���餤�j�����Τ�O<%=ren%>,�������B�i��<%=jin%>�j�v��!! <%end if%></FONT></FORM></TD></TR></TABLE><p align="center"> 
<a href="#" onClick="MM_showHideLayers('Layer1','','show')"> </a></p><p align="center">&nbsp;</p><p align="center"><b><font color="#FFE3A5"> 
</font></b></p> 
</body> 
</html> 
<% 
rst.close 
cn.close 
set rst=nothing 
set cn=nothing 
%> 
