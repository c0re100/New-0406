<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name1=Request.form("name1")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name2")
name2=trim(name2)
name2=server.HTMLEncode(name2)
pass=request.form("pass")
pass=trim(pass)
liyou=Request.form("liyou")
liyou=trim(liyou)
liyou=server.HTMLEncode(liyou)
if name1="" or name2="" then
	Response.Write "<script Language=Javascript>alert('提示：姓名為空，不能操作!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
if pass="" then
	Response.Write "<script Language=Javascript>alert('提示：姓名為空，不能操作!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
if name1=name2 then
	Response.Write "<script Language=Javascript>alert('提示：取自己!');location.href = 'LIFEN.htm';</script>"	
	response.end
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
rs.open "SELECT 配偶,銀兩 FROM 用戶 WHERE id=" & info(9) & " and 密碼='" & pass & "'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：用戶名密碼不正確!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
xb1=trim(rs("配偶"))
if xb1<>name2 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示："& name2 &"也不是你的配偶！!');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
if rs("銀兩")<101000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：需要101000兩纔可以離婚的！!');location.href = 'LIFEN.htm';</script>"
	response.end
end if
rs.close
rs.open "SELECT ID FROM 用戶 WHERE trim(姓名)='" & name2 & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&name2&"在江湖上找不到，請找管理員解決！！');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-1000 where id="&info(9)
rs.close
rs.open "SELECT 姓名1 FROM 離婚 WHERE 姓名1='" & name1 & "' or 姓名2='" & name2 & "'",conn
If not(Rs.Bof OR Rs.Eof) Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：["&info(0)&"你在江湖上申請過離婚了！！');location.href = 'LIFEN.htm';</script>"	
	response.end
end if
sz = "'" & name1 & "','" & name2 & "','" & liyou & "', now()"
conn.Execute "INSERT INTO 離婚 (姓名1, 姓名2, 理由, 時間) VALUES(" & sz & ")"
conn.execute "delete * from 月老 where now()-時間>10"
Response.Write "<script Language=Javascript>alert('提示："& info(0) &"你的離婚登記成功！！');location.href = 'LIFEN.htm';</script>"	
response.end
%>
 