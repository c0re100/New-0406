<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if request.form("submit")<>"還款" then
	dk=request.form("dk")
	if dk="" then Response.Redirect "daikuan.asp"
	if not isnumeric(dk) then response.redirect"../error.asp?id=464"
	dk=int(dk)
	if dk<1 then dk=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT allvalue FROM 用戶 WHERE id="&info(9),conn
If Rs.Bof OR Rs.Eof Then 
	set rs=nothing
	rs.Close
	conn.Close
	set conn=nothing
	Response.Redirect "../error.asp?id=440"
end if
	allvalue=rs("allvalue")
	bigdk=allvalue*100
	if dk-bigdk>0 then 
		set rs=nothing
		rs.Close
		conn.Close
		set conn=nothing
		Response.Redirect "../error.asp?id=482"
	end if
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
sj=n & "-" & y & "-" & r
rs.close
rs.open "select id,還貸記錄 from 貸款 where 貸款人='"&info(0)&"'",conn
	'貸款操作項
	if rs.BOF or rs.EOF then
		rs.close
		rs.open "select id from 貸款 where 還貸記錄=true",conn
			if rs.BOF or rs.EOF then
				conn.Execute "insert into 貸款(貸款人,貸款日期,貸款金額,還貸記錄)values('"&info(0)&"',"&sqlstr(sj)&","&dk&",false)"
			else
				id=rs("id")
				conn.execute "update 貸款 set 貸款日期="&sqlstr(sj)&",貸款金額="&dk&" ,貸款人='"&info(0)&"',還貸記錄=false where id="&id
			end if
	else
			if rs("還貸記錄")=true then
				'如果有記錄則程式將更新其記錄
				conn.execute "update 貸款 set 貸款日期="&sqlstr(sj)&",貸款金額="&dk&",還貸記錄=false where 貸款人='"&info(0)&"'"
			else
				Response.Redirect "../error.asp?id=483"
			end if
	end if
conn.execute "update 用戶 set 銀兩=銀兩+"&dk&" where id="&info(9)
rs.close
set rs=nothing
conn.Close
set conn=nothing
Response.Redirect "daikuan.asp"
'貸款操作項在此結束
	Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
	End Function

else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select 貸款金額 from 貸款 where 貸款人='"&info(0)&"'",conn
	dk=rs("貸款金額")
	rs.close
	rs.open "select 銀兩 from 用戶 where id="&info(9)
	qian=rs("銀兩")
	q=qian-(dk*1.5)
if q<0 then
	response.redirect "daikuan.asp"
	response.end
end if
conn.execute "update 貸款 set 還貸記錄=true where 貸款人='"&info(0)&"'"
conn.execute "update 用戶 set 銀兩=銀兩-"&dk*1.5&" where 姓名='"&info(0)&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "daikuan.asp"
end if
%>