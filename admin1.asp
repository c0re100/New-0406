<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="chkhk.asp"-->
<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("TWT_ARR_ArgALL")="" then response.end
TWT_ArrArg=split(session("TWT_ARR_ArgALL"),"=")
name=TWT_ArrArg(0)
grade=TWT_ArrArg(2)
myid=TWT_ArrArg(1)
a=request("a")
id=request("id")
name2=request.form("name2")
if InStr(name2,"=")<>0 or InStr(name2,"`")<>0 or InStr(name2,"'")<>0 or InStr(name2," ")<>0 or InStr(name2," ")<>0 or InStr(name2,"'")<>0 or InStr(name2,chr(34))<>0 or InStr(name2,"\")<>0 or InStr(name2,",")<>0 or InStr(name2,"<")<>0 or InStr(name2,">")<>0 then Response.Redirect "error.asp?id=120"
if InStr(name2,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "error.asp?id=120"
if InStr(active,"=")<>0 or InStr(active,"`")<>0 or InStr(active,"'")<>0 or InStr(active," ")<>0 or InStr(active," ")<>0 or InStr(active,"'")<>0 or InStr(active,chr(34))<>0 or InStr(active,"\")<>0 or InStr(active,",")<>0 or InStr(active,"<")<>0 or InStr(active,">")<>0 then Response.Redirect "error.asp?id=120"
userip=Request.ServerVariables("REMOTE_ADDR")
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr

Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function

sub setlogo(thislogo)
sql="INSERT INTO logdata (logtime,name,ip,opertion) VALUES ("
sql=sql & SqlStr(sj) & ","
sql=sql & SqlStr(name) & ","
sql=sql & SqlStr(userip) & ","
sql=sql & SqlStr(thislogo) & ")"
conn.Execute sql
end sub

if a="a" then
	sql="SELECT * FROM �Τ� WHERE �m�W='" & name2 & "'"
else
	sql="SELECT * FROM �Τ� WHERE ID=" & id
end if

set rs=conn.execute(sql)
if rs.eof or rs.bof then
	mess=name2 & "���O���򤤤H�I"
else
	thisname=rs("�m�W")
	thisgrade=rs("grade")

	sql="select * from �Τ� where �m�W='" & name & "' and grade>=9 and ����='������'"
	set rs=conn.execute(sql)

	if rs.eof or rs.bof or grade<10 then
		mess=name & " _ �Q�ťH�U�L�v��x���i��޲z"
	elseif int(thisgrade)>=int(grade) then
		mess=name & " _ ���޲z���Ÿ�A�ۦP�Ϊ̤�A���A�A�L�v�i��޲z"
	else
        	shefeng=rs("����")
		select case a
		case "a"
			sql="update �Τ� set grade=6, ����='������',����='�L' where �m�W='" & name2 & "'"
			conn.execute sql
			mess="�A���\�a��[" & thisname & "]�۸u�����������u�@�H��"
		case "b"
			if shefeng<>"�x��" then
			mess="�A���F���ѴN�Q�ư����w�ڡH<br>�������������x���ׯ�}���x���H��"
			else
			sql="update �Τ� set ����='�L', ����='�L',grade=4 where id=" & id
			conn.execute sql
			mess="�A���\�a��[" & thisname & "]�q�������}���F�I"
			end if
		case "c"
			if grade<10 then 
			mess=" - �Q�ťH�U�L�v�i��ɭ��žާ@�C"
			elseif int(sf)>=10 and shefeng<>"�x��" then
			mess=" - �A���O�������x���A�L�v���ɧO�H���Q�ů����C"
			else
			sf=request("sf")
			if int(sf)>=10 and shefeng<>"�x��" then sf=9
			sql="update �Τ� set grade='" & sf & "' where id=" & id
			conn.execute sql
			mess="�A���\�a��[" & thisname & "]�������q" & thisgrade & "�Žհʬ�"&sf&"�ŤF�I"
			end if
		end select
		call setlogo(mess)
	end if
end if
conn.close
set conn=nothing
%>
<style>td{font-size:14}</style>
<body bgcolor=#000000 background="pic/bg.jpg">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr>
<td bgcolor=#FFFFFF height="89"> 
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=admin.asp>��^</a></p>
</td></tr>
</table>
</td>
</tr>
</table>
<div align="center"><br>
<font size="2" color="#FFFFFF"><b>CNET����� (C) 2001-2002 </b></font></div>
</body>