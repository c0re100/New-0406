<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%><%


%>
<!--#include file="dbpath.asp"-->
<!--#include file="char.inc"-->
<!--#include file="set.asp"-->
<%
city=request.form("city")
barname=request.form("barname")
typical=request.form("typical")
number=request.form("number")
add=request.form("add")
people=request.form("people")
pwd=request.form("pwd")
pwd2=request.form("pwd2")
phone=request.form("phone")
email=request.form("email")
zip=request.form("zip")
ip=request.form("ip")
homepage=request.form("homepage")
intros=request.form("intros")
bad=split(badstr,"|")
	for i=0 to UBound(bad)
	people=Replace(people,bad(i),"**")
	next
bad=split(badstr,"|")
	for i=0 to UBound(bad)
	barname=Replace(barname,bad(i),"**")
	next
bad=split(badstr,"|")
	for i=0 to UBound(bad)
	intros=Replace(intros,bad(i),"**")
	next

'**********檢查是否填寫了所有項，如果不是側自動返回申請頁面
if city="" or barname="" or typical="" or add="" or pwd="" or people="" or intros="" then
errmsg=errmsg & "請填寫完整網吧信息！\n"
end if
if pwd<>pwd2 then
errmsg=errmsg & "兩次輸入的密碼不相同！\n"
end if

dim rsc,errmsg
Set rsc = Conn.Execute("select * from bar where barname = '" & barname & "'")
if not rsc.eof then errmsg=errmsg & "網吧名稱已被注冊，請改名！\n"

if errmsg<>"" then
    Conn.Close
    Set conn = nothing
    Set rsc = nothing
    response.write("<script>alert('" & errmsg & "');history.go(-1)</script>")
    response.end
end if
'**********檢查結束**********

set rs=server.createobject("adodb.recordset")
sql="select * from bar where (id is null)" 

rs.open sql,conn,1,3
rs.addnew
rs("city")=htmlencode(city)
rs("barname")=htmlencode(barname)
rs("typical")=htmlencode(typical)
rs("number")=htmlencode(number)
rs("add")=htmlencode(add)
rs("people")=htmlencode(people)
rs("pwd")=pwd
rs("phone")=htmlencode(phone)
rs("email")=htmlencode(email)
rs("zip")=htmlencode(zip)
rs("date")=htmlencode(date())
rs("ip")=htmlencode(ip)
rs("homepage")=htmlencode(homepage)
rs("intros")=htmlencode(intros)

rs.update
rs.close
conn.close
set conn=nothing
Response.Redirect "index.asp?type=regok"%>