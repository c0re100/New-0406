<%'卡片
function kapian(fn1,to1,toname)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
if toname="大家"  then
	Response.Write "<script language=JavaScript>{alert('用卡對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(0)=toname and instr(";解除卡;陷害卡;查稅卡;暗黑卡;欺騙卡;搶親卡;喫豆卡;搗亂卡;殘害卡;踢人卡;逮捕卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己進行操作！');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
fn1=trim(fn1)
rs.open "select 物品名 FROM 物品 WHERE 類型='卡片'  and 數量>0 and 物品名='" & fn1 & "' and 擁有者="& info(9),conn
if rs.eof or rs.bof  then
	Response.Write "<script language=JavaScript>{alert('你並沒有["&fn1&"]這種卡片,若想再次使用,請您購買！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

select case fn1
case "暗黑卡"
conn.Execute ("update 用戶 set kill=0 where 姓名='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('成功對["&toname&"]使用了暗黑卡！');}</script>"
kapian=info(0)&"對["&toname&"]使用了暗黑卡片，使對方的被殺次數為0,繼續受害了..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "解除卡"
conn.Execute ("update 用戶 set 保護=false,操作時間=now() where 姓名='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('成功解決了["&toname&"]的練功狀態！');}</script>"
kapian=info(0)&"偷偷使用了解除卡片，也不知道哪一位大蝦要倒霉了..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "欺騙卡"
rs.close
rs.open "select 存款 FROM 用戶 WHERE 姓名='"&toname&"'",conn
yin=rs("存款")
yin=yin/2
conn.Execute ("update 用戶 set 存款=存款+" & yin & "  where id="&info(9))
conn.Execute ("update 用戶 set 存款=存款-" & yin & "  where 姓名='" & toname & "'")
Response.Write "<script language=JavaScript>{alert('你成功的把對方身上的存款搶走了1/2！');}</script>"
rs.close
rs.open "select 存款 FROM 用戶 WHERE 姓名='" & toname & "'",conn
if rs("存款")<0 then 
conn.Execute ("update 用戶 set 存款=0 where 姓名='" & toname & "'")
end if
kapian=info(0)&"對["&toname&"]使用了欺騙卡，成功的把對方身上的存款騙走1/2,共計"&yin&"..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "搶親卡"
rs.close
rs.open "select 配偶 FROM 用戶 WHERE id="&info(9),conn
yin=rs("配偶")
if yin<>"無"  then
	Response.Write "<script language=JavaScript>{alert('你有老婆了呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
rs.open "select 配偶 FROM 用戶 WHERE 姓名='"&toname&"'",conn
yina=rs("配偶")
if yina="無" then
	Response.Write "<script language=JavaScript>{alert('對方還沒老婆呀，你要搶什麼呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
	conn.Execute ("update 用戶 set 配偶='" & yina & "'  where id="&info(9))
    conn.Execute ("update 用戶 set 配偶='無'  where 姓名='" & toname & "'")
conn.Execute ("update 用戶 set 配偶='" & info(0) & "'  where 姓名='" & yina & "'")
	Response.Write "<script language=JavaScript>{alert('你成功的把對方的老婆搶了過來，哈哈！');}</script>"
kapian=info(0)&"對["&toname&"]使用了搶親卡，成功的把["&toname&"]的老婆["&yina&"]給搶了過來..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "閃電卡"
	conn.Execute ("update 用戶 set 內力=內力/2  where 姓名='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('你成功的使用了閃電卡！');}</script>"
kapian=info(0)&"對"&toname&"使用了閃電卡，使對方的內力減半..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "查稅卡"
	conn.Execute ("update 用戶 set 存款=存款*0.7 where 姓名='" & toname & "'")
    conn.Execute ("update 用戶 set 銀兩=銀兩*0.7 where 姓名='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('成功的對對方使用查稅卡！');}</script>"
kapian=info(0)&"對"&toname&"使用了查稅卡，使對方身上的銀兩和銀行存款被收取30%稅款..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")

case "搶奪卡"
rs.close
rs.open "select 姓名,銀兩 FROM 用戶 WHERE 姓名='"&toname&"'",conn
yin=int(rs("銀兩")*0.5)
conn.Execute ("update 用戶 set 銀兩=銀兩+" & yin & "  where id="&info(9))
conn.Execute ("update 用戶 set 銀兩=銀兩-" & yin & "  where 姓名='" & toname & "'")
Response.Write "<script language=JavaScript>{alert('你成功的把對方身上的銀子搶走了1/2！');}</script>"
rs.close
rs.open "select 銀兩 FROM 用戶 WHERE 姓名='" & toname & "'",conn
if rs("銀兩")<0 then 
conn.Execute ("update 用戶 set 銀兩=0  where 姓名='" & toname & "'")
end if
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
kapian=info(0)&"對"&toname&"使用了搶奪卡，成功的把對方身上的銀子搶走了3/4,共計"&yin&"..."
case "陷害卡"
conn.Execute ("update 用戶 set 內力=int(內力/3),體力=int(體力/3) where 姓名='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('["&toname&"]的體力體力隻剩原來的1/3了！');}</script>"
kapian=info(0)&"終於忍不可忍，拿出自己陷害卡，向"&toname&"的頭上打去,"&toname&"隻覺眼前一黑，體力、內力損失大半.."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "法術卡"
conn.Execute ("update 用戶 set 法力=法力+2000 where id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的法力增加2000點！');}</script>"
kapian=info(0)&"使用了法術卡，這下可爽了，法力一時之間增加了2000點！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "怪風卡"
conn.Execute ("update 用戶 set 保護=false,殺人數=0,操作時間=now() where 會員等級=1")
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功使用怪風卡！');}</script>"
kapian="<center><font size=4 color=red>【"&info(0)&"偷偷使用了怪風卡，1級會員的保護都解除了...】</center><bgsound src=wav/003.MID loop=1></font>"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")

case "補血卡"
	conn.Execute ("update 用戶 set 體力=體力+50000,內力=內力+50000  where id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的體力、體力都增加了5萬點！');}</script>"
kapian=info(0)&"使用了補血卡，這一下可是有福了，體力、內力暴漲5萬點！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "破產卡"
			conn.Execute ("update 用戶 set 銀兩=0  where 姓名='" & toname & "'")
			Response.Write "<script language=JavaScript>{alert('你成功的使用了破產卡！');}</script>"
kapian=info(0)&"對"&toname&"使用了破產卡，使其身上的銀兩一分都沒有了，好慘哦..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "漲錢卡"
conn.Execute ("update 用戶 set 存款=存款+88880000  where  id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的存款上漲了8888萬！');}</script>"
kapian=info(0)&"使用了漲錢卡<img src=card/mhk.jpg>，自己的小荷包都裝不下了，8888萬呀！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "練功卡"
conn.Execute ("update 用戶 set 武功=武功+10000  where  id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]使用了練功卡，武功上漲1萬！');}</script>"
kapian=info(0)&"使用了練功卡，武功可是大幅度上漲，看來江湖又要不太平了！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "老虎卡"
	conn.Execute ("update 用戶 set 體力=體力/2  where 姓名='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('你成功的使用了老虎卡！');}</script>"
kapian=info(0)&"對"&toname&"使用了老虎卡，使對方的體力減半..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "復位卡"
conn.Execute ("update 用戶 set 殺人數=0  where  id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的殺人數恢復到初始狀態，又可以開始殺人了！');}</script>"
kapian=info(0)&"使用了復位卡，使自己的殺人數恢復到初始狀態，又可以開始殺人了呀！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "加點卡"
	conn.Execute ("update 用戶 set allvalue=allvalue+1000  where id="&info(9))
Response.Write "<script language=JavaScript>{alert('["&info(0)&"]使用了加點卡，泡點數上漲1000點！');}</script>"
kapian=info(0)&"使用了加點卡，呵。。不用泡點也加點，真是有福氣！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "財神卡"
yin2=100000*Application("ljjh_chatrs"&session("nowinroom"))
conn.Execute ("update 用戶 set 銀兩=銀兩+" & yin2 & " where 姓名='" & toname & "'")
Response.Write "<script language=JavaScript>{alert('你成功的使用了財神卡！');}</script>"
kapian=info(0)&"對"&toname&"使用了財神卡，使其收到聊天室在線人數乘10萬兩的銀子..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "衰神卡"
rs.close
rs.open "select 銀兩 FROM 用戶 WHERE 姓名='" & toname & "'",conn
yin3=rs("銀兩")/Application("ljjh_chatrs"&session("nowinroom"))
	useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
			online=Split(Trim(useronlinename)," ",-1)
			x=UBound(online)
			for i=0 to x
conn.Execute ("update 用戶 set 銀兩=銀兩+" & yin3 & " where 姓名='" & online(i) & "'")
			next
conn.Execute ("update 用戶 set 銀兩=0 where 姓名='" & toname & "'")
			Response.Write "<script language=JavaScript>{alert('你成功的使用了衰神卡！');}</script>"
kapian=info(0)&"對"&toname&"使用了衰神卡，聊天室在座的每個人分得"&toname&"的"&yin3&"兩銀子，看他衰成那樣，哈哈哈哈~~~！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "喫豆卡"
	conn.Execute ("update 用戶 set 暴豆時間=now()  where 姓名='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('對["&toname&"]使用了喫豆卡，他不能再暴豆了！');}</script>"
kapian=info(0)&"實在對"&toname&"的行為看不過去，使用了喫豆卡，"&toname&"大叫一聲，暈死過去。豆豆沒有了..."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")

case "暴豆卡"
	conn.Execute ("update 用戶 set 暴豆時間=now()  where id="&info(9))
Response.Write "<script language=JavaScript>{alert('對["&info(0)&"]使用了暴豆卡，現在暴豆成功了！');}</script>"
kapian=info(0)&"大叫著，我暴，我暴....從口袋中拿出暴豆卡，暴豆成功！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "道德卡"
	conn.Execute ("update 用戶 set 道德=道德+100000  where id="&info(9))
rs.close
rs.open "select 等級,道德,道德加 FROM 用戶 WHERE id="&info(9),conn
ddj=(rs("等級")*1440+2200)+Rs("道德加")
if rs("道德")>ddj then
conn.Execute ("update 用戶 set 道德=" & ddj & "  where id="&info(9))
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的道德上漲了10萬！');}</script>"
        kapian=info(0)&"使用了道德卡，自己的道德不斷上漲，漲了10萬哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "搗亂卡"
	conn.Execute ("update 用戶 set 武功=int(武功/3)  where 姓名='"&toname&"'")
Response.Write "<script language=JavaScript>{alert('對["&toname&"]使用了搗亂卡，他武功隻剩1/3了！');}</script>"
kapian=info(0)&"大叫著，誰也不要攔著我，我要為民除害！使用了搗亂卡，"&toname&"的武功失去大半...."
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "魅力卡"
	conn.Execute ("update 用戶 set 魅力=魅力+200000  where id="&info(9))
rs.close
rs.open "select 等級,魅力,魅力加 FROM 用戶 WHERE id="&info(9),conn
mlj=(rs("等級")*1120+2100)+rs("魅力加")
if rs("魅力")>mlj then
conn.Execute ("update 用戶 set 魅力=" & mlj & "  where id="&info(9))
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的魅力上漲了20萬！');}</script>"
        kapian=info(0)&"使用了魅力卡，自己的魅力不斷上漲，漲了20萬哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")

case "悔悟卡"
rs.close
rs.open "select 道德 FROM 用戶 WHERE id="&info(9),conn
if rs("道德")<0 then
conn.Execute ("update 用戶 set 道德=10000  where id="&info(9))
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的道德恢復正常！');}</script>"
        kapian=info(0)&"使用了悔悟卡，自己的道德恢復了正常，可以不做惡人了哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "幸運卡"
randomize timer
r=int(rnd*1200)+800
rs.close
rs.open "select 金幣 FROM 用戶 WHERE id="&info(9),conn
conn.Execute ("update 用戶 set 金幣=金幣+"&r&"  where id="&info(9))
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]的幸氣真好，金幣增加了"&r&"個！');}</script>"
        kapian=info(0)&"使用了幸運卡，金幣突然增加了"&r&"個，幸氣還不差哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "殘害卡"
rs.close
rs.open "select 體力 FROM 用戶 WHERE  姓名='" & toname & "'",conn
conn.Execute ("update 用戶 set 體力=1000  where 姓名='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功的使用了殘害卡！');}</script>"
        kapian=info(0)&"對" & toname & "使用了殘害卡，" & toname & "被天外神兵一頓狠揍，體力僅剩1000了哦,諸位大俠此時再不下手更待何時哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "非禮卡"
rs.close
rs.open "select 體力 FROM 用戶 WHERE  id="&info(9),conn
conn.Execute ("update 用戶 set 魅力=魅力-1000  where 姓名='" & toname & "'")
conn.Execute ("update 用戶 set 魅力=魅力+10000,知質=知質+10000  where id="&info(9))
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功的使用了非禮卡！');}</script>"
        kapian=info(0)&"對" & toname & "使用了非禮卡，" & toname & "被"&info(0)&"當眾調戲，一時氣暈過去，魅力減少1000，"&info(0)&"哦，你太可惡之極了餓,"&info(0)&"道德減少1000點，魅力、資質增加10000點！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "加豆卡"
rs.close
rs.open "select 泡豆點數 FROM 用戶 WHERE  id="&info(9),conn
conn.Execute ("update 用戶 set 泡豆點數=泡豆點數+10000  where id="&info(9))
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功的使用了加豆卡！');}</script>"
        kapian=info(0)&"對自己使用了加豆卡，立時口袋裡豆豆多多，樂樂多多哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case "貞節卡"
	Response.Write "<script language=JavaScript>{alert('貞節卡無需使用，隻需保留！');}</script>"
Response.End
case "赦免卡"
	Response.Write "<script language=JavaScript>{alert('赦免卡無需使用，隻需要保留！');}</script>"
Response.End
case "踢人卡"
rs.close
rs.open "select grade FROM 用戶 WHERE  姓名='" & toname & "'",conn
if rs("grade")>5 then
	Response.Write "<script language=JavaScript>{alert('失敗，你不能對高級管理員操作或該人是特別保護對像!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功的使用了踢人卡！');}</script>"
        kapian=info(0)&"對" & toname & "使用了踢人卡，隻見" & toname & "一個筋鬥倒飛出了聊天室哦！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
call boot(toname)
case "逮捕卡"
rs.close
rs.open "select grade FROM 用戶 WHERE  姓名='" & toname & "'",conn
if rs("grade")>5 then
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
conn.Execute ("update 用戶 set 狀態='獄',登錄=now()+3  where 姓名='" & toname & "'")
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功的使用了逮捕卡！');}</script>"
        kapian=info(0)&"對" & toname & "使用了逮捕卡，隻見" & toname & "頭上冒出一隻黑手一抓就把" & toname & "抓到了天牢！"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
call boot(toname)
case "變性卡"
rs.close
rs.open "select 性別,配偶,二婚 FROM 用戶 WHERE  id="&info(9),conn
xb=rs("性別")
pei=rs("配偶")
er=rs("二婚")
if xb="男" then
xb="女"
else
xb="男"
end if
conn.Execute ("update 用戶 set 性別='"&xb&"',配偶='無',二婚='無'  where id="&info(9))
conn.Execute ("update 用戶 set 配偶='無'  where 姓名='"&pei&"'")
conn.Execute ("update 用戶 set 二婚='無'  where 姓名='"&er&"'")
	Response.Write "<script language=JavaScript>{alert('["&info(0)&"]成功的使用了變性卡！');}</script>"
        kapian=info(0)&"對自己使用了變性卡，隻見"&info(0)&"大叫：我變，我變，我變變變,"&info(0)&"成功的把自己改造成一個???"
conn.Execute ("update 物品 set 數量=數量-1 where 擁有者="&info(9) &" and 物品名='"&fn1&"'")
case else
	Response.Write "<script language=JavaScript>{alert('你並沒有["&fn1&"]這種卡片,若想再次使用,請您購買！');}</script>"
	Response.End
end select
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','操作','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>