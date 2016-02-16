<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
username=info(0)
grade=info(3)
if grade<5 then
	Response.Write "<script language=JavaScript>{alert('江湖小輩，本怡紅院不要沒名望的姑娘!');location.href = 'index.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 性別,魅力 FROM 用戶 WHERE 姓名='"&username&"'"
set Rs=conn.Execute(sql)
sex=rs("性別")
meimao=rs("魅力")
if sex<>"女" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你有沒有搞錯呀，怡紅院裡可不要男的哦!');location.href = 'index.asp';}</script>"
	response.end
end if
if meimao<1000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你有沒有搞錯呀，你長的這麼丑還來這兒!');location.href = 'index.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from 名妓 where 姓名='" & username & "'"
set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
sql="insert into 名妓(姓名,美貌度,介紹) values ('" & username & "'," & meimao & ",'無')"
conn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT id FROM 用戶 WHERE 姓名='"&username&"'"
set Rs=conn.Execute(sql)
sql="update 用戶 set 銀兩=銀兩+1000000 where 姓名='"& username &"'"
conn.Execute sql
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('恭喜，你正式成為本院的姑娘，得到賣身銀兩1000000萬！!');location.href = 'index.asp';}</script>"
	response.end
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經是本怡紅院的姑娘了，怎麼還來登記呀！!');location.href = 'index.asp';}</script>"
	response.end
connt.close
end if
%> 