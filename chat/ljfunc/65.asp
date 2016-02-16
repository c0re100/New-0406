<%'老鴇
function laobao(to1,toname)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 等級 FROM 用戶 WHERE id="&info(9),conn
grade1=rs("等級")
rs.close
rs.open "select 性別,魅力,等級,門派  FROM 用戶 WHERE 姓名='"&toname&"'",conn
meimao=rs("魅力") 
xingbie=rs("性別") 
grade2=rs("等級")
meipai=rs("門派")
if grade1<garde2 then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('人家江湖經驗比你豐富多了，沒賣了你已經不錯了！');}</script>" 
Response.End 
end if 
if meipai="官府" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('你不能對官府人員進行此操作！');}</script>" 
Response.End 
end if 
if xingbie="男" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>{alert('你不能對男人操作！');}</script>" 
Response.End 
end if 

'rs.close 
'%><!--#include file="jiu.asp"--><%
'sql="select * from 名妓 where 姓名='"&toname&"'"
'set rs=connt.execute(sql)
'if not (rs.eof and rs.bof) then 
'Response.Write "<script language=javascript>{alert('人家已經是小姐了！');}</script>" 
'rs.close 
'set rs=nothing 
'conn.close 
'set conn=nothing 
'Response.End 
'end if 
'sql="insert into 名妓(姓名,美貌度) values ('" & toname & "'," & meimao & ")"
'			connt.execute sql
'rs.close
'rs.open "select id FROM 用戶 WHERE id="&info(9),conn
'conn.execute "update 用戶 set 銀兩=銀兩+1000000,道德=道德-1000 where id="&info(9)
laobao=info(0)& "的嘴皮子真是好用，" 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>