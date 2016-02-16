<%
Response.Buffer=true
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
message="姓名不能為空"
elseif pass="" then
message="密碼不能為空"
elseif mess<0 or mess>1 then
message="財產分割數目不對"
else
%><!--#include file="../data.asp"--><%
'校驗用戶 魅力大於100，錢大於1000
sql="SELECT * FROM 用戶 WHERE 姓名='" & name1 & "' and 密碼='" & pass & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
response.write "不能登記！請檢查你的姓名和密碼"
conn.close
response.end
else
xb1=trim(rs("性別"))
if rs("銀兩")>=1000 then
rs.close
set rs=nothing
set rs=conn.execute("SELECT * FROM 月老 WHERE trim(姓名1)='" & name1 & "'trim(姓名2)='" & name2 & "'")
if not (rs.bof or rs.eof) then

sql="update 用戶 set 銀兩=銀兩-5000 where 姓名='" & name1& "'"
conn.execute sql
sz = "'" & name1 & "','" & name1& "','" & mess & "', now()"
into_db = "INSERT INTO 月老 (姓名1, 姓名2, 說明, 時間) VALUES(" & sz & ")"
conn.Execute(into_db)
sql="delete * from 月老 where now()-時間>10"
conn.execute sql
message="登記成功！收取費用5000兩！"
else %>
<script language=vbscript>
MsgBox "你沒有老婆啊？搞錯了吧！"
location.href = "../jh.asp"
</script>

<%               end if
else %>
<script language=vbscript>
MsgBox "你哪裡有5000兩啊，沒錢還想離婚？做夢去吧！"
location.href = "../jh.asp"
</script>

<%		end if
end if
conn.close
end if
%>
<script language=vbscript>
MsgBox "<%=message%>"
location.href = "../jh.asp"
</script>
 