<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if request.form("submit")<>"�ٴ�" then
	dk=request.form("dk")
	if dk="" then Response.Redirect "daikuan.asp"
	if not isnumeric(dk) then response.redirect"../error.asp?id=464"
	dk=int(dk)
	if dk<1 then dk=1
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT allvalue FROM �Τ� WHERE id="&info(9),conn
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
rs.open "select id,�ٶU�O�� from �U�� where �U�ڤH='"&info(0)&"'",conn
	'�U�ھާ@��
	if rs.BOF or rs.EOF then
		rs.close
		rs.open "select id from �U�� where �ٶU�O��=true",conn
			if rs.BOF or rs.EOF then
				conn.Execute "insert into �U��(�U�ڤH,�U�ڤ��,�U�ڪ��B,�ٶU�O��)values('"&info(0)&"',"&sqlstr(sj)&","&dk&",false)"
			else
				id=rs("id")
				conn.execute "update �U�� set �U�ڤ��="&sqlstr(sj)&",�U�ڪ��B="&dk&" ,�U�ڤH='"&info(0)&"',�ٶU�O��=false where id="&id
			end if
	else
			if rs("�ٶU�O��")=true then
				'�p�G���O���h�{���N��s��O��
				conn.execute "update �U�� set �U�ڤ��="&sqlstr(sj)&",�U�ڪ��B="&dk&",�ٶU�O��=false where �U�ڤH='"&info(0)&"'"
			else
				Response.Redirect "../error.asp?id=483"
			end if
	end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&dk&" where id="&info(9)
rs.close
set rs=nothing
conn.Close
set conn=nothing
Response.Redirect "daikuan.asp"
'�U�ھާ@���b������
	Function SqlStr(data)
	SqlStr="'" & Replace(data,"'","''") & "'"
	End Function

else
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select �U�ڪ��B from �U�� where �U�ڤH='"&info(0)&"'",conn
	dk=rs("�U�ڪ��B")
	rs.close
	rs.open "select �Ȩ� from �Τ� where id="&info(9)
	qian=rs("�Ȩ�")
	q=qian-(dk*1.5)
if q<0 then
	response.redirect "daikuan.asp"
	response.end
end if
conn.execute "update �U�� set �ٶU�O��=true where �U�ڤH='"&info(0)&"'"
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-"&dk*1.5&" where �m�W='"&info(0)&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "daikuan.asp"
end if
%>