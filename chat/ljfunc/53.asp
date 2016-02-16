<%'老鴇 
function laobao(to1,toname) 
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('拐賣對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("ljjh_usermdb") 
rs.open "select 等級 from 用戶 where id="&info(9),conn 
grade1=rs("等級")
rs.close 
rs.open "select id,性別,魅力,等級,門派 from 用戶 where 姓名='"&toname&"'",conn 
idd=rs("id")
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
rs.close
rs.open "select 物品名 FROM 物品 WHERE 類型='卡片'  and 數量>0 and 物品名='貞節卡' and 擁有者="& idd,conn
if not (rs.eof or rs.bof)  then

conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&idd&" and 物品名='貞節卡'")
	Response.Write "<script language=JavaScript>{alert('對方身上現有貞節卡'，你此次操作失敗！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
Response.Write "<script language=javascript>{alert('你不能對男人操作！');}</script>" 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
Response.End 
end if 
rs.close 
rs.open "select 姓名 FROM 名妓 WHERE 姓名='"&toname&"'",conn 
if not (rs.eof and rs.bof) then 
Response.Write "<script language=javascript>{alert('人家已經是小姐了！');}</script>" 
rs.close 
set rs=nothing 
conn.close 
set conn=nothing 
Response.End 
end if 
rs.close
set rs=nothing
conn.execute "update 用戶 set 銀兩=銀兩+500000,職業='采冰',道德=道德-1000 where id="&info(9) 
conn.execute "insert into 名妓(姓名,美貌度,介紹) values ('" & toname & "'," & meimao & ",'"&info(0)&"')" 
laobao=info(0) & "的嘴皮子真是好用，連蒙帶騙的就把"& toname &"賣到怡紅院了，得到好處費50萬，道德下降1000." 
conn.close
set conn=nothing 
end function 
%>