<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"
Response.Buffer=true
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

getpass=trim(Request.Form("getpass"))
username=trim(Request.Form("username"))
email=trim(Request.Form("email"))
if getpass="" or username="" or email="" then 
Response.Write "<script language=JavaScript>{alert('�藍�_�A��J�����~�I�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>"
Response.End 
end if
if InStr(getpass,"=")<>0 or InStr(getpass,"`")<>0 or InStr(getpass,"'")<>0 or InStr(getpass," ")<>0 or InStr(getpass," ")<>0 or InStr(getpass,"'")<>0 or InStr(getpass,chr(34))<>0 or InStr(getpass,"\")<>0 or InStr(getpass,",")<>0 or InStr(getpass,"<")<>0 or InStr(getpass,">")<>0 then 
Response.Write "<script language=JavaScript>{alert('�藍�_�A��J�����~�I�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>"
Response.End 
end if
if InStr(username,"=")<>0 or InStr(username,"`")<>0 or InStr(username,"'")<>0 or InStr(username," ")<>0 or InStr(username," ")<>0 or InStr(username,"'")<>0 or InStr(username,chr(34))<>0 or InStr(username,"\")<>0 or InStr(username,",")<>0 or InStr(username,"<")<>0 or InStr(username,">")<>0 then 
Response.Write "<script language=JavaScript>{alert('�藍�_�A��J�����~�I�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>"
Response.End 
end if
if InStr(email,"=")<>0 or InStr(email,"`")<>0 or InStr(email,"'")<>0 or InStr(email," ")<>0 or InStr(email," ")<>0 or InStr(email,"'")<>0 or InStr(email,chr(34))<>0 or InStr(email,"\")<>0 or InStr(email,",")<>0 or InStr(email,"<")<>0 or InStr(email,">")<>0 then 
Response.Write "<script language=JavaScript>{alert('�藍�_�A��J�����~�I�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>"
Response.End 
end if
if len(getpass)=0 or len(username)=0 or len(email)=0 then
Response.Write "<script language=JavaScript>{alert('�藍�_�A��J�����~�I�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>"
Response.End 
end if
temppass=StrReverse(left(getpass&"godxtll,./",10))
templen=len(getpass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
getpass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.Open ("select �m�W from �Τ� where �K��='"&getpass&"' and �m�W='"&username&"' and �H�c='"&email&"'"),conn
if rs.EOF or rs.BOF then
Response.Write "<script language=JavaScript>{alert('�藍�_�A��J�����~�I�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>"
else
conn.Execute ("update �Τ� set �K�X='s`le555556' where �K��='"&getpass&"' and �m�W='"&username&"' and �H�c='"&email&"' and �K��='"&getpass&"'")
%>
<title>��^�K�X�I</title>
<body bgcolor="#990000" TEXT="#FFFFCC">
<div align="center">�A���K�X�O�Q�]��<FONT COLOR="#00FF00">123456</FONT>�A�кɧ�<FONT COLOR="#0000FF"><a href="#"
onClick="window.open('../yamen/modify.asp','modify','scrollbars=yes,resizable=yes,width=400,height=350')" title="�ק�K�X" target="_self">�ק�</a></FONT>�A���s�K�X 
<br> <input type=button value=�������f onClick='window.close()' name="button"> </div><%
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing

%>