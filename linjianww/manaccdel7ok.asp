<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT �m�W FROM �Τ� WHERE  �|������<1 and times=1 and CDate(lasttime)<now()-29"
rs.open sql,conn,1,1
if rs.Eof and rs.Bof then
rs.close
conn.close
set rs=nothing
set conn=nothing
Response.Redirect "manerr.asp?id=131"
end if
totalrec=rs.RecordCount
rs.close
set rs=nothing
sql="DELETE * FROM �Τ� WHERE times=1 and CDate(lasttime)<now()-29"
conn.Execute(sql)
conn.close
set conn=nothing%><html>
<head>
<title>�b���޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">�i�b���޲z�j</font></h1>
<font color="#FF0000">�i�M�z30�Ѱ��ιL�������b���j</font></div>
<hr noshade size="1" color=009900>
<b> �ާ@���� </b><br>
�@�M�����ιL�����A�B�W�L30�Ѫ��b�� <font color=red><%=totalrec%></font> �ӡC
<div align=center><a href="manacc.asp">��^</a></div>
</body>
</html>
 