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
nowinroom=session("nowinroom")
if Instr(LCase(Application("ljjh_useronlinename"&nowinroom)),LCase(" "&info(0)&" "))=0 then 
	Session("ljjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 

chatroomsn=session("nowinroom")

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
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
id=trim(request.querystring("id"))
fromname=trim(request.querystring("from1"))
to1=trim(request.querystring("to1"))
yn=trim(request.querystring("yn"))
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select �m�W FROM �Τ� WHERE id="&fromname,conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('�藍�_�A��誺�W�r���s�b�C');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
fromname=rs("�m�W")
rs.close
if to1="�j�a" or to1=fromname then 
response.write "<script Language='Javascript'>alert('�z�{�b�S���ѻP�P���A�ЦV�A�����ϥ�\""/�����J 1000 \""�R�O�n�D�}���C');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if info(0)<>to1 then
  msg="�H�a�S���ܽЧA�r�I"
  abc=0
elseif yn=0 then
  msg="�A�ڵ��H�a�����F�I"
  abc=1
  duidu="�i�ڵ��j["&info(0)&"]�ڵ�["& fromname &"]�������J�P�ШD!"
  duidu=duidu & "<br>.........�����J���d�o��A���٭n��s�@�U..."
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dpk.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dpk where ufrom='"& info(0) & "' or ufrom='"& fromname & "'"
Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
unpkname=rs("ufrom")
rs.close
conn.close
set rs=nothing
set conn=nothing
response.write "<script Language='Javascript'>alert('" & unpkname & "�٦��P���S�������A����A���}��');location.href = 'javascript:history.go(-1)';</script>"
abc=1
duidu="�i�X���j["&info(0)&"]���౵��["& fromname &"]�����ШD!,�]��["&unpkname&"]�٦���L���P���S�������C"
msg=unpkname & "�٦��P���S�������A����A���}���I"
else

sql="select duz from dpk where uto='"& info(0) & "$�U�`' and ufrom='"& fromname & "$�U�`' order by id Desc"
Set Rs=conn.Execute(sql)
	If rs.eof and rs.bof Then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('���~:�S���o�ӽ䧽!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
	end If

xiazhu=Rs("duz")
sql="delete * from dpk where uto='"& info(0) & "$�U�`' or ufrom='"& fromname & "$�U�`' or ufrom='"& info(0) & "$�U�`' or uto='"& fromname & "$�U�`'"
Set Rs=conn.Execute(sql)

for i= 1 to 18
Randomize
apk = Int(14 * Rnd+1)
Randomize
bbb= Int(4 * Rnd+1)
if bbb=1 then bbb2="��"
if bbb=2 then bbb2="��"
if bbb=3 then bbb2="��"
if bbb=4 then bbb2="��"
dpk=bbb2 & apk
dpk=convpk(dpk)
if instr(frpk,dpk & "|")=0 then
frpk=frpk & dpk & "|"
else 
i=i-1
end if
next
for i= 1 to 18
Randomize
apk = Int(14 * Rnd+1)
Randomize
bbb= Int(4 * Rnd+1)
if bbb=1 then bbb2="��"
if bbb=2 then bbb2="��"
if bbb=3 then bbb2="��"
if bbb=4 then bbb2="��"
dpk=bbb2 & apk
dpk=convpk(dpk)
if instr(frpk,dpk & "|")=0 and instr(topk,dpk & "|")=0 then
topk=topk & dpk & "|"
else 
i=i-1
end if
next
sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& info(0) & "','"& fromname & "','"& topk & "',false," & xiazhu & ",90)"
Set Rs=conn.Execute(sql)
sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& fromname & "','"& info(0) & "','"& frpk & "',true," & xiazhu & ",90)"
Set Rs=conn.Execute(sql)
duidu="<font color=green>["&info(0)&"]</font>�P�N�����J,<font color=green>[�F�C�դh]</font>�}�l���P......<br>"
duidu=duidu & "<font color=green>["&info(0)&"]</font>��F18�i�P<img src='duju/dpk/1.GIF'>"
duidu=duidu & "<input type=button value='�~�P' onclick=javascript:window.open('duju/dpk-xp.asp?name="&info(0)&"','f3','width=380,height=210') style='background-color:#86A231;color:FFFFFF;border: 1 double' onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=dpkname><img src='duju/dpk/2.GIF'>"
duidu=duidu & " <font color=green>["
duidu=duidu & fromname&"]</font>��F18�i�P<img src='duju/dpk/2.GIF'><input type=button value='�~�P' onclick=javascript:window.open('duju/dpk-xp.asp?name="
duidu=duidu & fromname &"','f3','width=380,height=210')  style=background-color:#86A231;color:FFFFFF;border: 1 double  onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=dpkname><img src='duju/dpk/1.GIF'>"
duidu=duidu & "<br>.....�ھڵP���W�w�A��[" & info(0) & "]���o�P�C<br>"
abc=1 
msg="�A�P�N���襴���J�F�A�еy�ԡI[����դh]���b���A�̬~�P�I"
set conn=nothing
set rs=nothing
end if
end if
if abc=1 then
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
sd(195)=info(0)
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="��"
sd(199)="<font color=#FFC01F>�i�����J�j</font>:"& duidu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock



end if
%>
<html>
<head>
<style>
a:link {text-decoration:none; font-size:14;color:#000000}
a:hover {text-decoration:underline;font-size:14; color:#000000; background:#ffffff}
a:visited {text-decoration:none;font-size:14; color:#000000}
td {font-size:9pt; color:#ff0000; line-height:16pt}
</style>
<META http-equiv='content-type' content='text/html; charset=big5'>
<title>���</title>
</head>
<body bgcolor=#0099FF ><table>
<tr>
<td> <font color="red"> </font><br>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
<tr> 
<td height="115"> 
<table border="0" bgcolor="#009900" cellspacing="0" cellpadding="2" width="361">
<tr> 
<td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">�F�C����</font><font color="red" size=2> 
</font></td>
<td width="29" height="9"> 
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr> 
<td width="16"><b><a href="<%if id="200" then%>javascript:window.close();<%else%>javascript:history.go(-1)<%end if%>" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="��^"><font color="000000">��</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="359" cellpadding="4">
<tr> 
<td width="59" align="center" valign="top" height="29"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="278" height="29"> <font color="red" size=2> 
<%=msg%>
</font></td>
</tr>
<tr> 
<td colspan="2" align="center" valign="top" height="58"> 
<input type=button value='����' onClick="javascript:window.close()" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223">
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<%
function convpk(dpk)
if dpk="��14" then dpk="�j��"
if dpk="��14" then dpk="�j��"
if dpk="��14" then dpk="�p��"
if dpk="��14" then dpk="�p��"
dpk=replace(dpk,"13","K")
dpk=replace(dpk,"12","Q")
dpk=replace(dpk,"11","J")
dpk=replace(dpk,"10","�Q")
dpk=replace(dpk,"1","A")
dpk=replace(dpk,"�Q","10")
convpk=dpk
end function 
if to1=jiutian_dabusi then
%>
<script language="JavaScript">window.close();</script>
<%end if%>
