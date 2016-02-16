<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
wg=request("wg")
a=request("a")
wg1=request.form("wg1")
nei=request.form("nl")
if nei>=100000000 then nei=99999999
name=request.form("name")
pass=request.form("pass")
if instr(repass,"'")>0 or instr(pass,",")>0 or instr(nei,",")>0 or instr(wg,",")>0 or instr(wg1,",")>0 or instr(name,",")>0 then
	response.write "你好呀！黑客先生，這回不靈了吧？！"
	response.end
end if
if instr(wg,"or")<>0 or instr(pai,"or")<>0 or instr(wg1,"or")<>0 or instr(nl,"or")<>0 or instr(name,"or")<>0 or instr(pass,"or")<>0 then
	response.write "你好呀！黑客先生，這回不靈了吧？！"
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
rs.open "SELECT ID FROM 用戶 where id="&info(9) &" and 密碼='" & pass & "'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你的用戶密碼錯誤啊"
	response.end
end if
rs.close
rs.open "SELECT 二婚 FROM 用戶 where id="&info(9) &" and 二婚<>'無'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你沒有二婚，不能創建合體技。"
	response.end
end if
ff=rs("二婚")
rs.close
rs.open "update 用戶 set 銀兩=銀兩-10000 where id="&info(9),conn
if a="m" then
	if info(1)="男" then
		ff1=info(0)
	else
		ff1=ff
		ff=info(0)
	end if
	conn.Execute "update 合體技 set 合體技='" & wg1 & "', 內力=" & nei & " where 姓名男='" & ff1 & "' or 姓名女='" & ff & "'"
end if
if a="n" then
	if info(1)="男" then
		ff1=info(0)
	else
		ff1=ff
		ff=info(0)
	end if
	conn.Execute "insert into 合體技(合體技,姓名男,姓名女,內力) values ('" & wg1 & "','" & ff1 & "','" & ff & "','" & nei & "')"
end if
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "stunt.asp"
%>
 