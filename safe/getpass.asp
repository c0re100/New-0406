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
Response.Write "<script language=JavaScript>{alert('對不起，輸入項有誤！！！！\n     按確定返回！');parent.history.go(-1);}</script>"
Response.End 
end if
if InStr(getpass,"=")<>0 or InStr(getpass,"`")<>0 or InStr(getpass,"'")<>0 or InStr(getpass," ")<>0 or InStr(getpass," ")<>0 or InStr(getpass,"'")<>0 or InStr(getpass,chr(34))<>0 or InStr(getpass,"\")<>0 or InStr(getpass,",")<>0 or InStr(getpass,"<")<>0 or InStr(getpass,">")<>0 then 
Response.Write "<script language=JavaScript>{alert('對不起，輸入項有誤！！！！\n     按確定返回！');parent.history.go(-1);}</script>"
Response.End 
end if
if InStr(username,"=")<>0 or InStr(username,"`")<>0 or InStr(username,"'")<>0 or InStr(username," ")<>0 or InStr(username," ")<>0 or InStr(username,"'")<>0 or InStr(username,chr(34))<>0 or InStr(username,"\")<>0 or InStr(username,",")<>0 or InStr(username,"<")<>0 or InStr(username,">")<>0 then 
Response.Write "<script language=JavaScript>{alert('對不起，輸入項有誤！！！！\n     按確定返回！');parent.history.go(-1);}</script>"
Response.End 
end if
if InStr(email,"=")<>0 or InStr(email,"`")<>0 or InStr(email,"'")<>0 or InStr(email," ")<>0 or InStr(email," ")<>0 or InStr(email,"'")<>0 or InStr(email,chr(34))<>0 or InStr(email,"\")<>0 or InStr(email,",")<>0 or InStr(email,"<")<>0 or InStr(email,">")<>0 then 
Response.Write "<script language=JavaScript>{alert('對不起，輸入項有誤！！！！\n     按確定返回！');parent.history.go(-1);}</script>"
Response.End 
end if
if len(getpass)=0 or len(username)=0 or len(email)=0 then
Response.Write "<script language=JavaScript>{alert('對不起，輸入項有誤！！！！\n     按確定返回！');parent.history.go(-1);}</script>"
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
rs.Open ("select 姓名 from 用戶 where 密取='"&getpass&"' and 姓名='"&username&"' and 信箱='"&email&"'"),conn
if rs.EOF or rs.BOF then
Response.Write "<script language=JavaScript>{alert('對不起，輸入項有誤！！！！\n     按確定返回！');parent.history.go(-1);}</script>"
else
conn.Execute ("update 用戶 set 密碼='s`le555556' where 密取='"&getpass&"' and 姓名='"&username&"' and 信箱='"&email&"' and 密取='"&getpass&"'")
%>
<title>找回密碼！</title>
<body bgcolor="#990000" TEXT="#FFFFCC">
<div align="center">你的密碼是被設為<FONT COLOR="#00FF00">123456</FONT>，請盡快<FONT COLOR="#0000FF"><a href="#"
onClick="window.open('../yamen/modify.asp','modify','scrollbars=yes,resizable=yes,width=400,height=350')" title="修改密碼" target="_self">修改</a></FONT>你的新密碼 
<br> <input type=button value=關閉窗口 onClick='window.close()' name="button"> </div><%
end if
rs.Close
set rs=nothing
conn.close
set conn=nothing

%>