<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then  response.redirect "../error.asp?id=425"
shopname=Request.QueryString("shopname")
%><!--#include file="data.asp"--><%
sql="SELECT * FROM �ө� where �ө��W='"&shopname&"'"
set rs1=conn1.Execute (sql)
if  rs1.EOF or rs1.BOF then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("�S���Ӱө�!")
history.back(1)
</script>
<%Response.End 
end if
%>
<head>
<title>�ק�ө�</title>
</head>

<body bgcolor="#99CCFF" BACKGROUND="../bg.gif" TEXT="#99CC99">
<p align="center"><FONT COLOR="#800080" SIZE="6">�ק�ө�</FONT></p><FORM METHOD="POST" ACTION="overshopok.asp"> 
<DIV ALIGN="center"> <CENTER> <TABLE BORDER="1" WIDTH="70%" CELLSPACING="0" CELLPADDING="0" BORDERCOLORLIGHT="#000000" BORDERCOLORDARK="#FFFFFF"> 
<TR> <TD WIDTH="39%">�ө��W:</TD><TD WIDTH="66%"><INPUT TYPE="text" NAME="shopname" SIZE="10" STYLE="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1" VALUE="<%=rs("�ө��W")%>"></TD></TR><TR> 
<TD WIDTH="39%">��&nbsp; �D:</TD><TD WIDTH="66%"><INPUT TYPE="text" NAME="name" SIZE="10" STYLE="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1" VALUE="<%=rs("���D")%>"></TD></TR> 
<TR> <TD WIDTH="39%">I D:</TD><TD WIDTH="66%"><INPUT TYPE="text" NAME="ID" SIZE="10" STYLE="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1" VALUE="<%=rs("ID")%>"></TD></TR> 
<TR> <TD WIDTH="39%">��&nbsp; ��:</TD><TD WIDTH="66%"><INPUT TYPE="text" NAME="memo" SIZE="33" STYLE="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1" VALUE="<%=rs("����")%>"></TD></TR> 
<TR> <TD WIDTH="39%">�g�綵��:</TD><TD WIDTH="66%"><INPUT TYPE="text" NAME="shoptype" SIZE="10" STYLE="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1" VALUE="<%=rs("�ө�����")%>"></TD></TR> 
<TR> <TD WIDTH="39%">���:</TD><TD WIDTH="66%"><INPUT TYPE="text" NAME="shopmoney" SIZE="10" STYLE="background-color: #99CCFF; color: #000000; border-style: solid; border-width: 1" VALUE="<%=rs("�`�U���")%>"></TD></TR> 
<TR> <TD WIDTH="105%" COLSPAN="2"> <P ALIGN="center"><INPUT TYPE="submit" VALUE="��          ��" NAME="B1"></TD></TR> 
</TABLE></CENTER></DIV></FORM><DIV ALIGN="CENTER"></DIV><%set rs1=nothing
conn1.close
set conn1=nothing
%>