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
rst.open"select * from 中獎號碼  where  datediff('d','"&tempdate&"',彩票日期)=0",cn,1,1
if rst.bof then
RANDOMIZE
guessno=int((99999-10000+1)*rnd+10000)
'guessno=12345
cn.execute"update 中獎號碼 set 中獎號碼='"&guessno&"',彩票日期='"&tempdate&"'"
application("guess")=guessno
end if
if application("guess")="" then application("guess")=rst("中獎號碼")
	sql="SELECT * FROM 中獎號碼"
	Set Rst=cn.Execute(sql)
	gold=rst("累計金額")
%>
<html>
<head>
<title>靈劍彩票中心</title> <meta http-equiv="Content-Type" content="text/html; charset=big5"> 
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
<tr> <td> <p align="center"><font size="+1" color="#0000FF">彩票說明</font></p><p> <font size="2"> 此彩票每日晚12時開獎，中獎號完全由電腦探制，站長無法操作，本期不中，獎金將移到下一期。每次一人隻可購買一個號碼，購買號碼為5位數字。買一個號碼花100兩銀子，累計金將加1000兩。</font></p><p align="right"><font size="2">------靈劍江湖彩票中心</font></p><p align="center">[<a href="#" onClick="MM_showHideLayers('Layer1','','hide')">關閉此窗口</a>]</p></td></tr> 
</table></div><TABLE WIDTH="463" BORDER="0" ALIGN="CENTER"><TR><TD><TABLE BORDER=1 CELLSPACING=0 CELLPADDING=2 ALIGN="center" BORDERCOLORDARK="#FFFFFF" WIDTH="32%" HEIGHT="31"> 
<TR ALIGN="center"> <TD BGCOLOR="#669900" WIDTH="100%" HEIGHT="25"><IMG BORDER="0" SRC="xycpzx.jpg" WIDTH="463" HEIGHT="85"></TD></TR> 
</TABLE><P ALIGN="CENTER"><A HREF="#" ONCLICK="MM_showHideLayers('Layer1','','show')"><FONT COLOR="#FFFFFF">購 
買 說 明</FONT></A><BR></P><TABLE WIDTH="263" BORDER="1" CELLSPACING="0" CELLPADDING="3" ALIGN="center" BORDERCOLORLIGHT="#000000" BORDERCOLORDARK="#FFFFFF" BGCOLOR="#FFFFFF"> 
<TR> <TD WIDTH="116">下期獎金己達到</TD><TD WIDTH="129"><%=gold%></TD></TR> <TR> <TD WIDTH="116">本期中獎號碼是</TD><TD WIDTH="129"><%=application("guess")%></TD></TR> 
<TR> <%
sql="SELECT * FROM 購買者 where datediff('d','"&tempdate&"',購買日期)=-1  and 號碼="&application("guess")&""
Set Rst=cn.Execute(sql)
if rst.bof then
ren="無人中獎"
else
ren=rst("購買者")
'中獎人加銀兩
sql="select * from 購買者 where 購買者='"&ren&"' and 領獎=false"
set rst=cn.execute(sql)
	if not(rst.bof) then 
	sql="update 購買者 set 領獎=true"
	set rst=cn.execute(sql)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
	sql="update 用戶 set 銀兩=銀兩+"&gold&" where 姓名='"&ren&"'"
	set rs=conn.execute(sql)
'加入中獎名
sql="INSERT into 中獎人(中獎者,中獎金額,中獎號碼,中獎日期)values('"&ren&"',"&gold&","&application("guess")&",'"&tempdate&"')"
Set Rst=cn.Execute(sql)
'累計金清0
	sql="update 中獎號碼 set 累計金額=0"
	set rst=cn.execute(sql)
	end if
end if
	sql="select * from 中獎人 where 中獎號碼="&application("guess")&" and 中獎者='"&ren&"'"
	set rst=cn.execute(sql)
	if not(rst.bof) then
	ren=rst("中獎者")
	jin=rst("中獎金額")
	else
	ren="無人中獎"
	jin="無人中獎"
	end if
	%> <TR> <TD WIDTH="116">本期中獎人是</TD><TD WIDTH="129"><%=ren%></TD></TR> 
<TR> <TD WIDTH="116">本期中獎金額</TD><TD WIDTH="129"><%=jin%></TD></TR> </TABLE><FORM METHOD="post" ACTION="buyguess.asp"> 
<TABLE WIDTH="235" BORDER="1" CELLSPACING="0" CELLPADDING="3" ALIGN="center" BORDERCOLORLIGHT="#000000" BORDERCOLORDARK="#FFFFFF" BGCOLOR="#FFFFFF"> 
<TR> <TD WIDTH="99">購 買 者</TD><TD WIDTH="106"><%=info(0)%></TD></TR> <TR> <TD WIDTH="99">您的號碼</TD><TD WIDTH="106"> 
<%
	sql="select * from 購買者 where 購買者='"&info(0)&"' and datediff('d','"&tempdate&"',購買日期)=0"
	set rst=cn.execute(sql)
	if rst.bof then
	%> <INPUT TYPE="text" NAME="guessnumber" SIZE="10" MAXLENGTH="10"> </TD></TR> 
<TR> <TD COLSPAN="2"> <DIV ALIGN="center"> <INPUT TYPE="submit" NAME="Submit" VALUE="購買"> 
   <INPUT TYPE="reset" NAME="Submit2" VALUE="清空"> </DIV></TD></TR> <%
else
myguess=rst("號碼")
response.write myguess
end if
%> </TABLE><FONT COLOR="#FFFFFF" SIZE="2"> <%if ren="無人中獎" then%> 本期無人中獎，獎金將累計到下一期，希望下一期的中獎者就是您! 
<BR> 今天購買，說不準明天中大獎的就是你喲！ <%else%> 今日中大獎的用戶是<%=ren%>,中獎金額可有<%=jin%>大洋喲!! <%end if%></FONT></FORM></TD></TR></TABLE><p align="center"> 
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
