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
	sql="SELECT * FROM 用戶 WHERE 姓名='" & name2 & "'"
else
	sql="SELECT * FROM 用戶 WHERE ID=" & id
end if

set rs=conn.execute(sql)
if rs.eof or rs.bof then
	mess=name2 & "不是江湖中人！"
else
	thisname=rs("姓名")
	thisgrade=rs("grade")

	sql="select * from 用戶 where 姓名='" & name & "' and grade>=9 and 門派='六扇門'"
	set rs=conn.execute(sql)

	if rs.eof or rs.bof or grade<10 then
		mess=name & " _ 十級以下無權對官府進行管理"
	elseif int(thisgrade)>=int(grade) then
		mess=name & " _ 對方管理等級跟你相同或者比你高，你無權進行管理"
	else
        	shefeng=rs("身份")
		select case a
		case "a"
			sql="update 用戶 set grade=6, 門派='六扇門',身份='無' where 姓名='" & name2 & "'"
			conn.execute sql
			mess="你成功地把[" & thisname & "]招聘為六扇門的工作人員"
		case "b"
			if shefeng<>"掌門" then
			mess="你做了長老就想排除異已啊？<br>隻有六扇門的掌門纔能開除官府人員"
			else
			sql="update 用戶 set 身份='無', 門派='無',grade=4 where id=" & id
			conn.execute sql
			mess="你成功地把[" & thisname & "]從六扇門開除了！"
			end if
		case "c"
			if grade<10 then 
			mess=" - 十級以下無權進行升降級操作。"
			elseif int(sf)>=10 and shefeng<>"掌門" then
			mess=" - 你不是六扇門掌門，無權提升別人為十級站長。"
			else
			sf=request("sf")
			if int(sf)>=10 and shefeng<>"掌門" then sf=9
			sql="update 用戶 set grade='" & sf & "' where id=" & id
			conn.execute sql
			mess="你成功地把[" & thisname & "]的身份從" & thisgrade & "級調動為"&sf&"級了！"
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
<p align=center><a href=admin.asp>返回</a></p>
</td></tr>
</table>
</td>
</tr>
</table>
<div align="center"><br>
<font size="2" color="#FFFFFF"><b>CNET中文網 (C) 2001-2002 </b></font></div>
</body>