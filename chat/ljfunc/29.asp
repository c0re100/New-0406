<%'經驗
function jingyan(fn1,to1,toname)
if toname="大家" or toname=info(0) then
Response.Write "<script language=JavaScript>{alert('傳經驗對像有錯，請看仔細了！！');}</script>"
Response.End
exit function
end if
if info(3)<120 then
Response.Write "<script language=JavaScript>{alert('傳經驗需要江湖等級120級的才可以操作！');}</script>"
Response.End
end if
fn1=abs(fn1)
if fn1>500  then
if info(4)=0 then 
Response.Write "<script language=JavaScript>{alert('非會員傳經驗值不能超過500,會員不能超過10000');}</script>"
Response.End
else
if fn1>10000 then
Response.Write "<script language=JavaScript>{alert('非會員傳經驗值不能超過500,會員不能超過10000');}</script>"
Response.End
end if
end if
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select allvalue FROM 用戶 WHERE id="&info(9),conn
if rs("allvalue")<fn1 then
Response.Write "<script language=JavaScript>{alert('你沒有那麼多的經驗值無法傳送！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if

conn.execute "update 用戶 set allvalue=allvalue-" & fn1 & " where id="&info(9)
conn.execute "update 用戶 set allvalue=allvalue+" & fn1 & " where 姓名='"&toname&"'"
jingyan=info(0) & "把" & fn1 & "的經驗傳給了" & toname &",而自己的江湖等級降低了，大無畏呀！"
'記錄
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','傳授經驗','"& jingyan & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>