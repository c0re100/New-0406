<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<body topmargin="6" leftmargin="0" bgcolor="#FFFFFF" background="../../images/7.jpg">
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
sl=int(abs(Request.form("sl")))
if sl<1 or sl>9 then
	Response.Redirect "../../error.asp?id=72"
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
select case id

case 1
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if yin*sl<10000000 or jin*sl<1000 then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='子彈' and 擁有者=" &info(9),conn
If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('子彈',"&info(9)&",'彈藥',0,0,100,2012,2012,0,0,0,"&sl&",2000000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
	end if
case 2
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*20000000>yin or sl*5000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='老虎肉' and 擁有者=" &info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('老虎肉',"&info(9)&",'藥品',0,0,0,150,150,"&sl&",2000,400,400,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update 物品 set 銀兩=2000,數量=數量+1,會員="&aaao&" where  id="&id
end if
case 3
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*30000000>yin or sl*10000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='狼牙棒' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('狼牙棒',"&info(9)&",'法器',0,0,1000,2001,2001,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id 
	end if
case 4
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*40000000>yin or sl*20000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='破天錐' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('破天錐',"&info(9)&",'法器',0,0,1000,2002,2002,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 5
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*50000000>yin or sl*30000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='血滴子' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('血滴子',"&info(9)&",'法器',0,0,1000,2004,2004,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 6
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*60000000>yin or sl*40000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='搶劫令' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('搶劫令',"&info(9)&",'法器',0,0,1000,2005,2005,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 7
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*70000000>yin or sl*50000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='紅寶石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('紅寶石',"&info(9)&",'法器',0,0,1000,2006,2006,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 8
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*80000000>yin or sl*60000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='綠寶石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('綠寶石',"&info(9)&",'法器',0,0,1000,2007,2007,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 9
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*90000000>yin or sl*70000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='藍寶石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('藍寶石',"&info(9)&",'法器',0,0,1000,2008,2008,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 10
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*100000000>yin or sl*80000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='魔力鑽石' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('魔力鑽石',"&info(9)&",'法器',0,0,100,1100,1100,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if

case 11
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*200000000>yin or sl*90000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='生日蛋糕' and 擁有者=" &info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('生日蛋糕',"&info(9)&",'法器',0,0,100,2009,2009,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id 
	end if
case 12
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*300000000>yin or sl*100000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='絕情刀' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('絕情刀',"&info(9)&",'法器',0,0,1000,2010,2010,0,0,0,"&sl&",200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id
	end if
case 13
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*400000000>yin or sl*110000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='水晶球' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('水晶球',"&info(9)&",'法寶',100,0,0,0,0,"&sl&",200000,1600,1600,0,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 銀兩=200000,數量=數量+1,會員="&aaao&" where id="& id
	end if
case 14
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*500000000>yin or sl*120000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='大鯊魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大鯊魚',"&info(9)&",'暗器',0,0,100,300,300,0,-1200,-800,"&sl&",20000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=20000,數量=數量+1,會員="&aaao&" where id="& id
	end if
case 15
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*600000000>yin or sl*130000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='大草魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('大草魚',"&info(9)&",'藥品',0,0,100,301,301,0,300,50,"&sl&",2000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=2000,數量=數量+1,會員="&aaao&" where id="&id
	end if
case 16
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*700000000>yin or sl*140000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='小鯉魚' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('小鯉魚',"&info(9)&",'藥品',0,0,100,302,302,0,100,30,"&sl&",500,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 銀兩=500,數量=數量+1,會員="&aaao&" where id="& id
	end if
case 17
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*800000000>yin or sl*150000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='排雲秘籍' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('排雲秘籍',"&info(9)&",'物品',0,0,0,0,0,"&sl&",200000,99004,99004,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id

	end if
case 18
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*900000000>yin or sl*160000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='特種金屬' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('特種金屬',"&info(9)&",'物品',0,0,0,0,0,"&sl&",10000,2003303,2003303,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id

	end if
case 19
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*100000000>yin or sl*170000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='明月秘訣' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('明月秘訣',"&info(9)&",'物品',0,0,0,0,0,"&sl&",200000,99003,99003,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id

	end if
case 20
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1100000000>yin or sl*180000>jin then 
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='浮影劍譜' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('浮影劍譜',"&info(9)&",'物品',0,0,0,0,0,"&sl&",200000,99005,99005,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id

	end if
case 21
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1200000000>yin or sl*190000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='雷霆秘籍' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
  conn.execute "insert into 物品(物品名,擁有者,類型,防御,體力,內力,堅固度,攻擊,說明,sm,數量,銀兩,等級,會員)  values('雷霆秘籍',"&info(9)&",'物品',0,0,0,40000,15000,99006,99006,"&sl&",10000000,0,"&aaao&")" 
else
id=rs("id")
conn.execute  "update 物品 set 數量=數量+1,會員="&aaao&"  where id="& id
end if
case 22
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1300000000>yin or sl*200000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='黑石' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
  conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,體力,內力,堅固度,說明,sm,數量,銀兩,等級,會員) values('黑石',"&info(9)&",'物品',0,0,0,0,0,2003305,2003305,"&sl&",10000,0,"&aaao&")" 
else
id=rs("id")
 conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id 
end if

case 23
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1400000000>yin or sl*210000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='金屬板' and 擁有者=" & info(9),conn
if rs.eof and rs.bof then
  conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,體力,內力,堅固度,說明,sm,數量,銀兩,等級,會員)  values('金屬板',"&info(9)&",'物品',0,0,0,0,0,2003304,2003304,"&sl&",10000,0,"&aaao&")" 
else
id=rs("id")
 conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id 
end if

case 24
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1500000000>yin or sl*220000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='朱石' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('朱石',"&info(9)&",'物品',0,0,0,2003306,2003306,0,0,0,"&sl&",10000,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
end if

case 25
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1600000000>yin or sl*230000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='冰水' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,sm,等級,會員)  values ('冰水',"&info(9)&",'藥品',0,0,100,151,151,"&sl&",2000,151,151,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
end if

case 26
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1700000000>yin or sl*240000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='鐵片' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('鐵片',"&info(9)&",'物品',0,0,0,2003301,2003301,0,0,0,"&sl&",10000,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
end if

case 27
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1800000000>yin or sl*250000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='礦石' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('礦石',"&info(9)&",'物品',0,0,0,2014,2014,0,0,0,"&sl&",800,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
end if

case 28
rs.open "select 銀兩,金幣 from 用戶 where id=" &info(9),conn
yin=rs("銀兩")
jin=rs("金幣")
if sl*1900000000>yin or sl*260000>jin then
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有那麼多錢和金幣！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩-" & yin*sl & ",金幣=金幣-"&jin*sl&" where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='木頭' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('木頭',"&info(9)&",'物品',0,0,100,2015,2015,0,0,0,"&sl&",800,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
end if
case else
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：操作錯誤！');location.href = 'javascript:history.go(-1)';</script>"
response.end
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：購買物品成功！');location.href = 'javascript:history.go(-1)';</script>"
response.end
%>
