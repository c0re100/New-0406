<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")

username=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 性別,銀兩 FROM 用戶 WHERE 姓名='"&username&"'"
set rs=conn.execute(sql)
sex=rs("性別")
yl=rs("銀兩")
if sex="男" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你有沒有搞錯呀，怡紅院裡可不要男的哦!');location.href = 'index.asp';}</script>"
	response.end
end if
if yl<5000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你有沒有搞錯呀，沒錢也想贖身!');location.href = 'index.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from 名妓 where 姓名='" & username & "'"
set rs=conn.execute(sql)
if not rs.eof then
sql="delete from 名妓 where 姓名='"& username &"'"
conn.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT id FROM 用戶 WHERE 姓名='"&username&"'"
set rs=conn.execute(sql)
sql="update 用戶 set 銀兩=銀兩-5000000 where 姓名='"& username &"'"
set rs=conn.execute(sql)
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('恭喜，你終於離開了青樓，還清貸款銀兩500萬!');location.href = 'index.asp';}</script>"
	response.end
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經不是本怡紅院的姑娘了，怎麼還來!');location.href = 'index.asp';}</script>"
	response.end
end if
%> 