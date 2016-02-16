<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name1=Request.form("name12")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name22")
name2=trim(name2)
name2=server.HTMLEncode(name2)
pass=request.form("pass2")
pass=trim(pass)
mess=Request.form("mess2")
mess=trim(mess)
mess=server.HTMLEncode(mess)
if name1="" or name2="" then
	Response.Redirect "../error.asp?id=433"
end if
if pass="" then	Response.Redirect "../error.asp?id=432"
if name1=name2 then
	Response.Redirect "../error.asp?id=434"
end if
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'校驗用戶 魅力大於100，錢大於1000
rs.open "SELECT 配偶,門派,性別,銀兩 FROM 用戶 WHERE id="&info(9)&" and 密碼='" & pass & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=431"
	response.end
end if
if rs("配偶")<>"無" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=430"
	response.end
end if
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('為了維護江湖管理、公平競爭,官府禁止結婚！');history.go(-1);</script>"
	response.end
end if
xb1=trim(rs("性別"))
if rs("銀兩")<1000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：錢不夠想取媳婦，沒門！！！');history.go(-1);</script>"
	response.end
end if
rs.close
rs.open "SELECT 門派,性別 FROM 用戶 WHERE trim(姓名)='" & name2 & "'",conn
if not (rs.bof or rs.eof) then
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('為了維護江湖管理、公平競爭,官府禁止結婚！');history.go(-1);</script>"
	response.end
end if
	xb2=trim(rs("性別"))
	if xb1<>xb2 then
		conn.execute "update 用戶 set 銀兩=銀兩-1000 where 姓名='" & name1& "'"
		sz = "'" & name1 & "','" & name2 & "','" & mess & "', now()"
		conn.Execute "INSERT INTO 月老 (姓名1, 姓名2, 說明, 時間) VALUES(" & sz & ")"
		conn.Execute "delete * from 月老 where now()-時間>5"
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Redirect "../ok.asp?id=600"
	else
		Response.Redirect "../error.asp?id=435" 
	end if
else
	Response.Write "<script Language=Javascript>alert('提示：沒有你要取的人呀！！');history.go(-1);</script>"
	response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%> 