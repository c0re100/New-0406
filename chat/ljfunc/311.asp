<%'單挑取消
function quxiaoa()
if info(2)<6 then
	Response.Write "<script language=JavaScript>{alert('你不是官府中人，請不要操作！');}</script>"
	Response.End
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 姓名1,姓名2 FROM 單挑",conn
xx1=rs("姓名1")
xx2=rs("姓名2")
Application.Lock
Application("diantiao")=""
Application.unLock
'conn.execute "delete * from 單挑 where 姓名1<>'無'"
conn.execute "update 單挑 set 姓名1='無',姓名2='無'"
quxiaoa="官府人員"&info(0)& "因參與挑戰的雙方["&xx1&"]----["&xx2&"]因為某些原因取消挑戰，其他人可以進行挑戰了  "

rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>