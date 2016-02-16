<%'逮捕
function daipu(fn1,to1,toname)
if info(2)<6 then
	Response.Write "<script language=JavaScript>{alert('你無權操作，你不是官府中人！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('逮捕對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名,grade,門派 from 用戶 where 姓名='"&toname&"'",conn
toname=rs("姓名")
if info(2)<=rs("grade") then
	Response.Write "<script language=JavaScript>{alert('失敗，你不能對高級管理員操作或該人是特別保護對像!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select id FROM 用戶 WHERE 姓名='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "select 物品名 FROM 物品 WHERE 類型='卡片'  and 數量>0 and 物品名='赦免卡' and 擁有者="& idd,conn
if not (rs.eof or rs.bof)  then

conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&idd&" and 物品名='赦免卡'")
	Response.Write "<script language=JavaScript>{alert('對方身上現有赦免卡，請給他次機會不要抓他餓，給站長個面子OK！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 狀態='獄',登錄=now()+3 where 姓名='"&toname&"'"
call boot(toname)
conn.execute "insert into 事務(管理捕快,管理對像,管理方式,行為時間,原因) values ('" & info(0) & "','" & toname & "','逮捕',now(),'" & fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
if info(5)<>"官府" then
daipu=info(0) & "[便衣聊管]:<font color=red><bgsound src=wav/oh_no.wav loop=1>官府的人拿了根鐵索套在" & toname & "的脖子上,連拉帶扯的把" & toname & "押到了牢房！[" & fn1 & "]</font>"
else
daipu=info(0) & "[官府人員]:<font color=red><bgsound src=wav/oh_no.wav loop=1>官府的人拿了根鐵索套在" & toname & "的脖子上,連拉帶扯的把" & toname & "押到了牢房！[" & fn1 & "]</font>"
end if
end function
%>