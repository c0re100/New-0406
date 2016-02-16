<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=trim(request.form("a"))
b=trim(request.form("b"))
c=trim(request.form("c"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM 用戶 where 姓名='"&a&"'",conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，沒有該用戶！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
select case b

case "赤劍"
rs.close
rs.open "select id from 物品 where 物品名='赤劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('赤劍',"&info(9)&",'右手',0,0,0,100000,100000,2003201,1,5,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "蝴蝶劍"
rs.close
rs.open "select id from 物品 where 物品名='蝴蝶劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('蝴蝶劍',"&info(9)&",'左手',100000,0,0,300000,0,2003202,1,15,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "松根劍"
rs.close
rs.open "select id from 物品 where 物品名='松根劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('松根劍',"&info(9)&",'右手',0,0,0,500000,200000,2003203,1,25,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "刺蠻劍"
rs.close
rs.open "select id from 物品 where 物品名='刺蠻劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('刺蠻劍',"&info(9)&",'左手',200000,0,0,800000,0,2003204,1,35,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "玄冥劍"
rs.close
rs.open "select id from 物品 where 物品名='玄冥劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('玄冥劍',"&info(9)&",'右手',0,0,0,1000000,300000,2003205,1,45,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "寒星魔劍"
rs.close
rs.open "select id from 物品 where 物品名='寒星魔劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('寒星魔劍',"&info(9)&",'左手',300000,0,0,1200000,0,2003206,1,55,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "行雲鬼刃"
rs.close
rs.open "select id from 物品 where 物品名='行雲鬼刃' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('行雲鬼刃',"&info(9)&",'右手',0,0,0,1400000,400000,2003207,1,65,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "嗜血劍"
rs.close
rs.open "select id from 物品 where 物品名='嗜血劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('嗜血劍',"&info(9)&",'左手',400000,0,0,1600000,0,2003208,1,75,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "青頭魔劍"
rs.close
rs.open "select id from 物品 where 物品名='青頭魔劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('青頭魔劍',"&info(9)&",'右手',0,0,0,1800000,500000,2003209,1,85,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "融雪劍"
rs.close
rs.open "select id from 物品 where 物品名='融雪劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('融雪劍',"&info(9)&",'左手',500000,0,0,2000000,0,20032010,1,95,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "蛟龍劍"
rs.close
rs.open "select id from 物品 where 物品名='蛟龍劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('蛟龍劍',"&info(9)&",'右手',0,0,0,2200000,600000,20032011,1,105,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "七星寶劍"
rs.close
rs.open "select id from 物品 where 物品名='七星寶劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('七星寶劍',"&info(9)&",'左手',600000,0,0,2400000,0,20032012,1,125,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "結拜巨刀"
rs.close
rs.open "select id from 物品 where 物品名='結拜巨刀' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('結拜巨刀',"&info(9)&",'右手',0,0,0,2600000,700000,20032013,1,155,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "鳳眼劍"
rs.close
rs.open "select id from 物品 where 物品名='鳳眼劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('鳳眼劍',"&info(9)&",'左手',700000,0,0,2800000,0,20032014,1,205,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "赤碧劍"
rs.close
rs.open "select id from 物品 where 物品名='赤碧劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('赤碧劍',"&info(9)&",'右手',0,0,0,3000000,800000,20032015,1,245,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "至尊劍"
rs.close
rs.open "select id from 物品 where 物品名='至尊劍' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values('至尊劍',"&info(9)&",'左手',800000,0,0,3200000,0,20032016,1,265,"&aaao&")")

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
case "屠龍刀"
rs.close
rs.open "select id from 物品 where 物品名='屠龍刀' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
conn.execute "insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,會員) values ('屠龍刀',"&info(9)&",'右手',0,0,0,3400000,900000,20032017,1,300,"&aaao&")"

else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+"&c&",會員="&aaao&" where id="&id
end if
end select
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','物品發送操作')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>
