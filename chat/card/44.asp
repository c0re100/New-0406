<%@ LANGUAGE=VBScript codepage ="950" %>

<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
'#####房間處理#####
zhbg_roominfo=split(Application("zhbg_room"),";")
chatinfo=split(zhbg_roominfo(nowinroom),"|")
if chatinfo(8)<>0 then
	Response.Write "<script language=JavaScript>{alert('提示：在["&chatinfo(0)&"]房間不可以使用卡片！');}</script>"
	Response.End
end if
if Instr(LCase(Application("zhbg_useronlinename"&nowinroom)),LCase(" "&zhbg_name&" "))=0 then 
	Session("zhbg_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
f=Minute(time())
zhbg_kptime=Application("zhbg_kptime")
if f>zhbg_kptime and chatinfo(0)<>"恩怨情仇" then
 Response.Write "<script language=javascript>{alert('提示：非恩怨房間用卡請在每小時的前"&zhbg_kptime&"分');}</script>"
 Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'對暫離開、點啞穴判斷
'call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&","&")
'對系統禁止字符處理
if zhbg_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=trim(mid(says,i+1))
if fnn1<>"歸來卡" then call dianzan(towho) 
says=kapian(mid(says,i+1),towho)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)

'卡片
function kapian(fn1,to1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
if zhbg_name=to1 and instr(";解除卡;陷害卡;查稅卡;吃豆卡;搗亂卡;逮捕卡;踢人卡;冬眠卡;搶親卡;情人卡;查稅卡;窮神卡;衰神卡;死神卡;分手卡;降級卡;怪獸卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能自己進行操作！');</script>"
	Response.End
end if
if to1="大家" and instr("逮捕卡;踢人卡;冬眠卡;搶親卡;情人卡;查稅卡;窮神卡;衰神卡;死神卡;分手卡;降級卡;怪獸卡;種子卡;種花卡;",fn1)<>0 then
	Response.Write "<script language=javascript>alert('【"&fn1&"】不能大家進行操作！');</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("zhbg_usermdb")
fn1=trim(fn1)
rs.open "select 轉生 from 用戶 where  姓名='"&to1&"'",conn,2,2
zstt=rs("轉生")
if zstt>=8 then
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##想對%%使用["&fn1&"]，但%%已經練成了超凡的能力(轉生大於等於8次)，此卡片已經被化解了...</font>"
exit function
end if
rs.close
rs.open "select 會員金卡,w5,門派 from 用戶 where  姓名='"&zhbg_name&"'",conn,2,2
mycard=abate(rs("w5"),fn1,1)
select case fn1
case "福神卡"
	conn.Execute ("update 用戶 set sl='福神',slsj=now()+3 where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]現在神靈賦身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]<img src=card/BXK.JPG><bgsound src=mav/XL.MID loop=1>，神啊，您與我同在...</font>"
case "財神卡"
	conn.Execute ("update 用戶 set sl='財神',slsj=now()+3 where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]現在神靈賦身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]<img src=card/caishen.JPG><bgsound src=mav/XL.MID loop=1>，神啊，您與我同在...</font>"
case "窮神卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set sl='窮神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]現在神靈與["&to1&"]賦身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]<img src=card/qiongshen.JPG><bgsound src=mav/XL.MID loop=1>，神啊，與你同在...</font>"
case "衰神卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set sl='衰神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]現在神靈與["&to1&"]賦身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]<img src=card/aishen.JPG><bgsound src=mav/XL.MID loop=1>，神啊，與你同在...</font>"
case "死神卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set sl='死神',slsj=now()+3 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]現在神靈與["&to1&"]賦身了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]<img src=card/sishen.JPG><bgsound src=mav/XL.MID loop=1>，神啊，與你同在...</font>"
case "送神卡"
	conn.Execute ("update 用戶 set sl='無',slsj=now() where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]成功解除了神靈賦身！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了["&fn1&"]<img src=card/BXK.JPG><bgsound src=mav/XL.MID loop=1>，神啊，88！see tomorrow!</font>"
case "色魔卡"
	conn.Execute ("update 用戶 set 結婚次數=結婚次數+5  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了色魔卡，也就表示你成了真正的色魔。！');}</script>"
	kapian="<font color=red>【色魔卡】<font color=" & saycolor & ">##使用了色魔卡<img src=card/semo.JPG><bgsound src=mav/XL.MID loop=1>，哈哈這下你那青梅竹馬，海誓山盟的愛人不要你了，結婚次數成5次了：呵呵，哭也沒辦法了，誰讓是個色鬼呢，哈哈。大家快打這個大色魔，快看看呀，才幾天，已經結過這麼多次婚了。哈哈。</font>"
case "非禮卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]<img src=card/feili.JPG><bgsound src=mav/XL.MID loop=1>...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用戶 set 魅力=魅力-500 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的魅力少了500！');}</script>"
	kapian="<font color=red>【卡片】<font color=" & saycolor & ">##色咪咪的，暈呼呼的就向%%的櫻桃小嘴一陣KISS,%%只覺一陣噁心。魅力少了500點..</font>"
case "合好卡"
	conn.Execute ("update 用戶 set 金幣=金幣+10 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]友誼之神,真誠的呼喚彼此,讓我們快樂，讓我們合好.！');}</script>"
	kapian="<img src=card/he1.gif><bgsound src=wav/hua.wav loop=1><font color=green>【卡片】<font color=" & saycolor & ">##真誠得對%%使用了合好卡<img src=card/hehao.gif>,為其送上了10個金幣.真是有心,朋友請不要一些小事,就失去了我們的友誼.願我們重歸於好,友誼地久天長,一起唱,朋友一生一起走,那些日子不在有,一句話,一輩子,一生情,一杯酒...</font>"
case "神劍卡"
	conn.Execute ("update 用戶 set 攻擊=攻擊+1000 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]本國度最高最美最強之物聖靈之劍,為對方獻上了1000攻擊,真是如意寶貝!');}</script>"
	kapian="<img src=card/jian.gif><bgsound src=readonly/okok.wav loop=1><font color=green>【卡片】<font color=" & saycolor & ">##真誠得對%%使用了依戀蝶最強最好最美之卡-------神劍卡<img src=card/jian1.gif>,夢幻之劍在瞬間發送耀眼的光彩.為其送上了1000攻擊,這回天下無敵了,%%真是太感動了啊.呵呵，可不可以在來一張呀:)...</font>"
case "魔盾卡"
	conn.Execute ("update 用戶 set 防禦=防禦+1000 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]非凡國度總站最高最美最強之物依戀聖龍之盾,為對方獻上了10000防禦!');}</script>"
	kapian="<img src=card/dun.gif><bgsound src=wav/hua.wav loop=2><font color=green>【卡片】<font color=" & saycolor & ">##真誠得對%%使用了非凡國度總站最強最好最美之卡-------魔盾卡<img src=card/dun1.gif>,夢幻之盾在瞬間發送耀眼的光彩.為其送上了1000防禦,這回天下無敵了,%%真是太感動了啊.呵呵，救世主呀:)...</font>"
case "天戒卡"
    conn.Execute ("update 用戶 set 法力=法力+20000 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]非凡國度總站最高最美最強之物聖龍之盾,為對方獻上了20000法力！');}</script>"
	kapian="<img src=card/jie.gif><bgsound src=wav/hua.wav loop=2><font color=green>【卡片】<font color=" & saycolor & ">##真誠得對%%使用了非凡國度總站最強最好最美之卡-------天戒卡<img src=card/jie1.gif>,夢幻之戒在瞬間發送耀眼的光彩.為其送上了2000法力,這回天下無敵了,%%真是太感動了啊.呵呵，救世主呀:)...</font>"
case "戰神卡"
    conn.Execute ("update 用戶 set 武功=武功+2000 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]非凡國度網最高最美最強之物聖蝶之刃,為對方獻上了2000武功！');}</script>"
	kapian="<img src=card/zhan.gif><bgsound src=cd.mid loop=1><font color=green>【卡片】<font color=" & saycolor & ">##真誠得對%%使用了非凡國度總站最強最好最美之卡-------戰神卡<img src=card/zhan1.gif>,夢幻之刃在瞬間發送耀眼的光彩.為其送上了2000武功,這回天下無敵了,%%真是太感動了啊.呵呵，救世主呀:)...</font>"
case "仙靴卡"
    conn.Execute ("update 用戶 set allvalue=allvalue+3000 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]非凡國度總站最高最美最強之物聖仙之靴,為對方獻上了30000存點！');}</script>"
	kapian="<img src=card/xue.gif><bgsound src=wav/hua.wav loop=2><font color=green>【卡片】<font color=" & saycolor & ">##真誠得對%%使用了非凡國度網最強最好最美之卡-------仙靴卡<img src=card/xue1.gif>,夢幻之靴在瞬間發送耀眼的光彩.為其送上了3000存點,%%快點存點哦:)...</font>"
case "天之卡"
    conn.Execute ("update 用戶 set allvalue=allvalue+1000") 
	Response.Write "<script language=JavaScript>{alert('["&to1&"]為非凡國度網所有玩家增加1000存點，夢凡對你表示忠心感謝！');}</script>"
	kapian="<img src=card/tzk.gif><bgsound src=wav/hua.wav loop=2><font color=green>【卡片】<font color=" & saycolor & ">##笑瞇瞇得對大家使用了天之卡<img src=card/tzk1.gif>,天地萬物,唯我獨尊.##的天之卡在明月的襯托下更加靚麗,所有玩家存點加1000.MYLOVE..這種寶貝,也真是##所擁有!讓大家都好好感謝你吧,##你真是非凡國度的福神啊.大家快點存點哦.</font>"
case "神之卡"
    conn.Execute ("update 用戶 set 金幣=金幣+88") 
	Response.Write "<script language=JavaScript>{alert('["&to1&"]為非凡國度總站所有玩家增加88個金幣！');}</script>"
	kapian="<img src=card/szk.gif><bgsound src=wav/hua.wav loop=2><font color=green>【卡片】<font color=" & saycolor & ">##笑瞇瞇得對大家使用了神之卡<img src=card/szk1.gif>,天地萬物,唯我獨尊.##的神之卡在明月的襯托下更加靚麗,所有玩家金幣加88.MYLOVE..這種寶貝,也真是##所擁有!讓大家都好好感謝你吧,##你真是非凡國度的福神啊.</font>"
case "魔之卡"
    conn.Execute ("update 用戶 set 狀態='正常',體力=體力+15000") 
	Response.Write "<script language=JavaScript>{alert('["&to1&"]天皇神卡真牛.！');}</script>"
	kapian="<img src=card/mzk.gif><bgsound src=wav/hua.wav loop=2><font color=green>【卡片】<font color=" & saycolor & ">##笑瞇瞇使用了如來佛主賜予的魔之卡<img src=card/mzk1.gif>,天地萬物,唯我獨尊.##的魔之卡在明月的襯托下更加靚麗,一道佛光在愛情大廳折射出神奇的光芒.那些冬眠的,被點穴的,被打死的,所有人都回來了,真是不知道趕怎樣形容這種寶貝了。##你真是非凡國度網的大救星啊.</font>"			
case "會員卡"
	conn.Execute ("update 用戶 set 會員等級=1,會員日期=date()+10  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]使用了最高最美最強之物會員之神卡,恭喜你成為非凡國度總站的會員啦！');}</script>"
	kapian="<font color=red>【卡片】<font color=" & saycolor & ">##使用了夢幻之卡，在瞬間發送耀眼的光彩.讓自己成為了一級會員，真是幸福啊,快刷新或重進吧！</font>"
case "超級紅卡"
 conn.Execute ("update 用戶 set 武功加=武功加+200,防禦加=防禦加+200,攻擊加=攻擊加+200,內力加=內力加+200,體力加=體力加+200 where 姓名='"&zhbg_name&"'")
 Response.Write "<script language=Javascript>{alert('對["&zhbg_name&"]使用了健身卡，鍛練成功！武功、內力、體力、防禦和攻擊上限值各漲200點！！！');}</script>"
 kapian="<font size=2 color=green>【超級卡片】<font color=" & saycolor & "><font color=red>##</font>大叫著，我要更強，我要更強....從口袋中拿出超級紅卡，經一翻艱苦訓練，武功、攻擊、防禦、內力、體力上限值各漲<font color=red>200</font>點！！！"
case "美麗卡"
	conn.Execute ("update 用戶 set 魅力=魅力+10000  where  姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]的魅力上漲了10000！');}</script>"
	kapian="<font color=red>【美麗卡】<font color=" & saycolor & ">##我要美麗，我要美麗~~我要變成最美麗的人，要賽過李嘉欣，劉德華，大家看這人呆呆的，出於同情，湊齊10000魅力給了她！</font>"
case "親親卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
    if rs("門派")="出家" then
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】你是出家人，搞錯了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用戶 where 姓名='"&to1&"'",conn
    if rs("門派")="出家" then
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】人家是出家人，搞錯了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing

    end if
    if rs("性別")=sex then
        kapian="<img src=card/qinqin.jpg><font color=green><bgsound src=003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用親親卡失敗:原因是"&to1&"與"&zhbg_name&"是同性!沒愛人了玻璃也想啊？</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.Execute ("update 用戶 set allvalue=allvalue+50  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]使用了親親卡，泡點數上漲50點！');}</script>"
    kapian="<img src=card/ai1.gif><font color=green><bgsound src=123.wav loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用親親卡<img src=card/qinqin.jpg>,終於如願以嘗的與"&to1&"在國度的大廳瘋狂的<img src=card/ai.gif>起來!真叫人羨慕啊！同時"&zhbg_name&"的泡點上升50點！你看！把"&zhbg_name&"樂得都跳起來了！大叫：<font color=red><b>抱抱卡在手帥哥美女別想走!</b></font></font>" 
case "羞羞卡"
    wnky=wnk(to1)
    if wnky<>"ok" then 
		kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
		exit function
    end if
	conn.Execute ("update 用戶 set 魅力=魅力+500 where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的魅力多了500！');}</script>"
	kapian="<font color=red>【卡片】<font color=" & saycolor & ">##臉泛桃花，含情默默地向自己追求已久的%%深情的吻去,%%只覺心神激盪，一種幸福感湧上心頭，彼此的心扉在那一刻盛開。感情也隨之昇華。魅力上升了500點..</font>"
case "抱抱卡"
 rs.close
 rs.open "select 性別 from 用戶 where 姓名='"&zhbg_name&"'",conn,2,2
 my_xb=rs("性別")
 rs.close
 rs.open "select 性別 from 用戶 where 姓名='"&to1&"'",conn,2,2
 to_xb=rs("性別")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('異性朋友你也抱？是不是同性戀啊？');}</script>"
	response.end
 end if
 conn.Execute ("update 用戶 set 體力=體力-100,內力=內力+500,魅力=魅力+500 where 姓名='"&to1&"'")
 conn.Execute ("update 用戶 set 體力=體力-100,內力=內力+500,道德=道德-100,allvalue=allvalue+300 where 姓名='"&zhbg_name&"'")
 kapian="<bgsound src=wav/baobao.wav loop=1><font color=green>【卡片】<font color=" & saycolor & ">##對%%使用了抱抱卡，終於如願以償的與%%在國度的大廳瘋狂的<img src=card/baobao.gif>起來！真叫人羨慕啊！同時##的經驗值增加了300點,你看，把##樂得三天三夜睡不著覺！大叫：<font color=red><b>抱抱卡在手帥哥美女別想走!</b></font>"
case "真愛卡"
 rs.close
 rs.open "select 性別 from 用戶 where 姓名='"&zhbg_name&"'",conn,2,2
 my_xb=rs("性別")
 rs.close
 rs.open "select 性別 from 用戶 where 姓名='"&to1&"'",conn,2,2
 to_xb=rs("性別")
 if my_xb=to_xb then
	Response.Write "<script language=JavaScript>{alert('異性朋友你愛什麼啊？是不是同性戀啊？');}</script>"
	response.end
 end if
 conn.Execute ("update 用戶 set 道德=道德+50,魅力=魅力+50,allvalue=allvalue+500 where 姓名='"&to1&"'")
 conn.Execute ("update 用戶 set 道德=道德+50,魅力=魅力+50,allvalue=allvalue-500 where 姓名='"&zhbg_name&"'")
 kapian="<bgsound src=wav/zhenai.wav loop=1><font color=green>【卡片】<font color=" & saycolor & ">##對%%使用了真愛卡，終於如願以償的與%%在國度的大廳瘋狂的<img src=card/zhenai.gif>給真愛的%%500經驗,愛是無私的奉獻，真叫國度人士眼紅啊！同時雙方的道德和魅力都上升50點，你看，把##樂得跳起來了，大叫：<font color=red><b>下輩子我還要你愛我可以嗎？</b></font>"
case "請神卡"
	dim sl(4)
	sl(0)="福神"
	sl(1)="財神"
	sl(2)="衰神"
	sl(3)="死神"
	sl(4)="窮神"
	randomize timer
	sss=int(rnd*4)+1
	if sss=5 then sss=4
	conn.Execute ("update 用戶 set sl='"&sl(sss)&"',slsj=now()+3 where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('成功使用了["&fn1&"]現在得到神靈["&sl(sss)&"]賦身！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了["&fn1&"]，神啊，您與我同在...請得大神["&sl(sss)&"]到自己的身上…</font>"
	erase sl
case "解除卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set 保護=false,操作時間=now() where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('成功解決了["&to1&"]的練功狀態！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷使用了解除卡片，也不知道哪一位大蝦要倒霉了...</font>"
case "陷害卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set 內力=int(內力/3),體力=int(體力/3) where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的體力體力只剩原來的1/3了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##終於忍不可忍，拿出自己陷害卡，向%%的頭上打去,%%只覺眼前一黑，體力、內力損失大半..</font>"
case "補血卡"
	conn.Execute ("update 用戶 set 體力=體力+50000,內力=內力+50000  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]的體力、體力都增加了5萬點！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了補血卡，這一下可是有福了，體力、內力暴漲5萬點！</font>"
case "漲錢卡"
	conn.Execute ("update 用戶 set 存款=存款+88880000  where  姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]的存款上漲了8888萬！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了漲錢卡，自己的小荷包都裝不下了，8888萬呀！</font>"
case "練功卡"
	conn.Execute ("update 用戶 set 武功=武功+10000  where  姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]使用了練功卡，武功上漲1萬！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了練功卡，武功可是大幅度上漲，看來國度又要不太平了！</font>"
case "攻防卡"
	conn.Execute ("update 用戶 set 攻擊=攻擊+300,防禦=防禦+300  where  姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]使用了攻防卡，攻擊和防禦各漲300點！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了攻防卡，攻擊和防禦大幅度上漲，看來國度又多出一位高手了！</font>"
case "加點卡"
	conn.Execute ("update 用戶 set allvalue=allvalue+1000  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('["&zhbg_name&"]使用了加點卡，泡點數上漲1000點！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了加點卡，呵。。不用泡點也加點，真是有福氣！</font>"
case "歸來卡"
wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
 if Instr(LCase(application("zhbg_zanli")),LCase("!"&to1&"!"))=0 then
  kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & sayscolor & ">"&zhbg_name&"對"&to1&"使用了歸來卡,可惜{"&to1&"}並不是在暫離狀態!白白浪費了一張歸來卡！</font>"
 else
  rs.close
  rs.open "select id,會員等級,等級,身份,門派,名單頭像,性別,好友名單,配偶,通緝 from 用戶 where 姓名='" & to1 &"'",conn,2,2
  jhid=rs("id")
  hydj=rs("會員等級")
  jhdj=rs("等級")
  jhsf=rs("身份")
  if rs("通緝")=True then
   jhmp="通緝犯"
  else
   jhmp=rs("門派")
  end if
  jhtx=rs("名單頭像")
  sex=rs("性別")
'  nowinroom=session("nowinroom")
  Application.Lock
  onlinelist=Application("zhbg_onlinelist"&nowinroom)
  onlinenum=UBound(onlinelist)
  for i=1 to onlinenum step 1
   onlinexx=split(onlinelist(i),"|")
   if onlinexx(0)=to1 then
   zhbg_zm=to1&"|"&sex&"|"&jhmp&"|"&jhsf&"|"&jhtx&"|"&jhdj&"|"&jhid&"|"&hydj&"|0"&"|"&onlinexx(9)
   onlinelist(i)=zhbg_zm
   exit for
   end if
  next
  Application("zhbg_onlinelist"&nowinroom)=onlinelist
  zhbg_zanli=split(application("zhbg_zanli"),";")
  for x=0 to UBound(zhbg_zanli)
   dxname=split(zhbg_zanli(x),"|")
   if dxname(0)="!"&to1&"!" then
    dxcz=zhbg_zanli(x)&";"
    application("zhbg_zanli")=replace(application("zhbg_zanli"),dxcz,"")
    kapian="<img src=card/youyi.jpg><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & sayscolor & ">"&zhbg_name&"對"&to1&"使用了歸來卡,成功地把{"&to1&"}從暫離狀態拖回來!</font>"
    exit for
   end if
  next
  Application.UnLock
 end if
case "吃豆卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set 暴豆時間=now()  where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('對["&to1&"]使用了吃豆卡，他不能再暴豆了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##實在對%%的行為看不過去，使用了吃豆卡，%%大叫一聲，暈死過去。豆豆沒有了...</font>"
case "暴豆卡"
	conn.Execute ("update 用戶 set 暴豆時間=now()  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('對["&zhbg_name&"]使用了暴豆卡，現在暴豆成功了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##:大叫著，我暴，我暴....從口袋中拿出暴豆卡，暴豆成功！</font>"
case "搗亂卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
	conn.Execute ("update 用戶 set 武功=int(武功/3)  where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('對["&to1&"]使用了搗亂卡，他武功只剩1/3了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##:大叫著，誰也不要攔著我，我要為民除害！使用了搗亂卡，"&to1&"的武功失去大半....</font>"
case "怪獸卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
        rs.close
        rs.open "SELECT * FROM 用戶 WHERE  姓名='" & to1 & "'",conn,2,2
        if rs("門派")="官府" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&",你不能對官府人員使用怪獸卡!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
        if rs("身份")="掌門" then 
            kapian="<img src=card/guaishou.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&",你不能對門派的掌門人<font color=red>[["&to1&"]]</font>使用怪獸卡!"
           	rs.close
		set rs=nothing
		conn.close
	    	set conn=nothing
	        exit function
	end if
	conn.Execute ("update 用戶 set 身份='殭屍王',狀態='殭屍王' where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('對["&to1&"]使用了怪獸卡，他已經變成了殭屍了！');}</script>"
	kapian="<img src=card/guaishou.jpg><font color=green>【卡片】<font color=" & saycolor & ">##:大叫著，"&to1&"變成了殭屍，請大家小心哦，如果怕的話請來我這邊，我可以保護大家....</font>"
case "清純卡"
	conn.Execute ("update 用戶 set 結婚次數=0  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了清純卡，結婚次數清0完成！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##好懷念過去的時光呀……</font>"
case "好人卡"
	conn.Execute ("update 用戶 set 總殺人=0  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了好人卡，殺人總數清0完成！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了[好人卡]，殺人總數為0了……</font>"
case "變性卡"
	rs.close
	rs.open "select 性別,配偶 FROM 用戶 WHERE 姓名='" & zhbg_name &"'",conn,2,2
	sex=rs("性別")
	pl=rs("配偶")
	rs.close
    if pl="無" then 
        if sex="男" then
               sql="update 用戶 set 性別='女' WHERE 姓名='" & zhbg_name & "'"
               xb="漂亮的女生"
         end if 
          if sex="女" then 
              sql="update 用戶 set 性別='男' WHERE 姓名='" & zhbg_name & "'"
              xb="英俊的男孩"
          end if
          Set Rs=conn.Execute(sql)  
            bianxi=zhbg_name&"使用變性卡後,終於如願以嘗,變成了"&xb&"!" 
        else
          bianxi="使用變性卡失敗!原因:"&zhbg_name&"是有家室的人呢,怎麼還想變性呀!這樣不是亂套了!為了懲戒像"&zhbg_name&"你這種道德敗壞的人,特此沒收"&zhbg_name&"的變性卡!"
        end if
	kapian="<img src=card/bxk.jpg></td><td><font color=green>【卡片】<font color=" & saycolor & ">##：這年頭，克隆技術真先進呀，我變，我變……(不是在變成克琳鈍，在作變性手術！）<br>【結果】"&bianxi&"</font>"
case "離婚卡"
	rs.close
	rs.open "select 配偶 FROM 用戶 WHERE 姓名='" & zhbg_name &"'",conn,2,2
	peiou=rs("配偶")
	if peiou="無" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】你有錢呀，根本沒有配偶呀！');</script>"
		Response.End
	end if
	conn.Execute ("update 用戶 set 配偶='無'  where 姓名='" & zhbg_name &"'")
	rs.close
	rs.open "select 配偶 FROM 用戶 WHERE 姓名='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update 用戶 set 配偶='無'  where 姓名='"&peiou&"'")
	end if
	rs.close
	kapian="<img src=card/lifen.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">##：想前想後，經過自己一番思想鬥爭，使用了離婚卡,終於想好了與["&peiou&"]離婚了……</font>"
case "搶親卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
    if rs("門派")="出家" then
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】你是出家人，搞錯了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用戶 where 姓名='"&to1&"'",conn,2,2
    if rs("門派")="出家" then
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】人家是出家人，搞錯了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("性別")
    if rs("配偶")<>"無" then
	    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用搶親卡失敗:原因是"&to1&"是有家室的人</font>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("門派")="官府" then 
           kapian="不可以對管理員使用搶親卡……"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from 用戶 where 姓名='"&zhbg_name&"'",conn,2,2
    if rs("配偶")<>"無" then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用搶親卡失敗:原因是自己已經是有家室的人了!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("性別")=sex then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用搶親卡失敗:原因是"&to1&"與"&zhbg_name&"是同性!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update 用戶 set 配偶='"&zhbg_name&"' where 姓名='"&to1&"'"
    conn.execute "update 用戶 set 配偶='"&to1&"' where 姓名='"&zhbg_name&"'"
    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用搶親卡,終於如願以嘗的與"&to1&"結為夫婦!</font>"
case "分手卡"
	rs.close
	rs.open "select 情人 FROM 用戶 WHERE 姓名='" & zhbg_name &"'",conn,2,2
	peiou=rs("情人")
	if peiou="無" then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】你有錢呀，根本沒有情人呀！');</script>"
		Response.End
	end if
	conn.Execute ("update 用戶 set 情人='無'  where 姓名='" & zhbg_name &"'")
	rs.close
	rs.open "select 情人 FROM 用戶 WHERE 姓名='"&peiou&"'",conn,2,2
	if not(Rs.Bof OR Rs.Eof) then
		conn.Execute ("update 用戶 set 情人='無'  where 姓名='"&peiou&"'")
	end if
	rs.close
	kapian="<img src=card/lifen.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】##：想前想後，經過自己一番思想鬥爭，使用了分手卡,終於想好了與["&peiou&"]分手了……</font>"
case "情人卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
    if rs("門派")="出家" then
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】你是出家人，搞錯了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
    rs.close
    rs.open "select * from 用戶 where 姓名='"&to1&"'",conn,2,2
    if rs("門派")="出家" then
		Response.Write "<script language=javascript>alert('【"&zhbg_name&"】人家是出家人，搞錯了！！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
	sex=rs("性別")
    if rs("情人")<>"無" then
	    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用情人卡失敗:原因是"&to1&"是有情人的人了</font>"
    	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
    	exit function
	end if
    if rs("門派")="官府" then 
           kapian="不可以對管理員使用情人卡……"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
    end if
    rs.close
    rs.open "select * from 用戶 where 姓名='"&zhbg_name&"'",conn,2,2
    if rs("情人")<>"無" then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用情人卡失敗:原因是自己已經是有情人的人了!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    if rs("性別")=sex then
        kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用情人卡失敗:原因是"&to1&"與"&zhbg_name&"是同性!</font>"
       	rs.close
		set rs=nothing
		conn.close
	    set conn=nothing
        exit function
    end if
    conn.execute "update 用戶 set 情人='"&zhbg_name&"' where 姓名='"&to1&"'"
    conn.execute "update 用戶 set 情人='"&to1&"' where 姓名='"&zhbg_name&"'"
    kapian="<img src=card/qinren.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用情人卡,終於如願以嘗的與"&to1&"結為情人!</font>"
case "逮捕卡"
wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
        rs.close
        rs.open "SELECT * FROM 用戶 WHERE  姓名='" & to1 & "'",conn,2,2
        if rs("門派")="官府" then 
            kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&",你不能對官府人員使用逮捕令!"
           	rs.close
			set rs=nothing
			conn.close
	    	set conn=nothing
	        exit function
		end if
        kapian="<img src=card/xianhao.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"拿出國度特許的逮捕令,把"&to1&"給抓了!"
        mzky=mzk(to1)
        if mzky="ok" then   
           conn.execute "update 用戶 set 狀態='獄',登錄=now()+3 where 姓名='" & to1 & "'"
            call boot(to1,to1&"被"&zhbg_name&"使用了逮捕令")
        else
           kapian=kapian&mzky
        end if
case "踢人卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
      rs.close
      rs.open "SELECT * FROM 用戶 WHERE  姓名='" & to1 & "'",conn,2,2
      if rs("門派")="官府" then 
        kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&",你不能對官府人員使用踢人卡!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
	  end if
       mtky=mtk(to1)
       if mtky="ok" then   
	      kapian="<img src=card/dz04.gif></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"使用踢人卡,飛起一腳，結果把"&to1&"踢了出去!"
    	  call boot(to1,to1&"被"&zhbg_name&"使用了踢人卡")     
    	else
			kapian=mtky
    	end if
case "冬眠卡"      
  wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
  rs.close
  rs.open "SELECT * FROM 用戶 WHERE  姓名='" & to1 & "'",conn,2,2
  if rs("門派")="官府" then
     kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&",你不能對官府人員使用冬眠卡!"
       	rs.close
		set rs=nothing
		conn.close
	   	set conn=nothing
        exit function
  end if
   rs.close
   qxky=qxk(to1)
   if qxky="ok" then   
	   conn.execute "update 用戶 set 登錄=now()+12/(4*144),狀態='眠' where 姓名='" & to1 & "'"
	   kapian="<img src=card/shuimian.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用冬眠卡!使"&to1&"睡著了!"
	   call boot(to1,to1&"被"&zhbg_name&"使用了冬眠卡")
   else
   		kapian=qxky
   end if
case "查稅卡" 
  wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
  rs.close   
  rs.open "SELECT * FROM 用戶 WHERE  姓名='" & to1 & "'",conn,2,2
  if rs("門派")="官府" then
    kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&",你不能對官府人員使用查稅卡!"
   	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
    exit function
  end if
    yl=rs("銀兩")+rs("存款")
    if yl<=10000 then
    	kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用查稅卡失敗:原因"&to1&"身上的銀子小於10000兩!"
	   	rs.close
		set rs=nothing
		conn.close
		set conn=nothing
	    exit function
	end if
    mhky=mhk(to1)
    if mhky="ok" then   
      yl=int(yl*0.02)
      conn.execute "update 用戶 set 銀兩=銀兩+"&yl&" where 姓名='"&zhbg_name&"'"
      if rs("銀兩")>=rs("存款") then
        conn.execute "update 用戶 set 銀兩=銀兩-"&yl&" where 姓名='"&to1&"'"
      else
        conn.execute "update 用戶 set 存款=存款-"&yl&" where 姓名='"&to1&"'"
      end if
      kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"使用查稅卡,查得"&to1&"共"&yl&"兩銀子,全部歸"&zhbg_name&"所有!"
    else
       kapian="<img src=card/chashui.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"對"&to1&"使用查稅卡失敗!"
       kapian=kapian&mhky
    end if
case "升級卡"
 rs.close
 rs.open "select allvalue,會員等級,等級 from 用戶 where 姓名='"&zhbg_name&"'",conn,2,2
 if rs("會員等級")>3 or rs("等級")>100 then
		Response.Write "<script language=javascript>alert('提示：100級以上或4級以上會員無需使用升級卡！');</script>"
		Response.End
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
    end if
 jhjy=rs("allvalue")
 jhdj=int(sqr(jhjy/50))
 jhadd=((jhdj+1)*(jhdj+1)-jhdj*jhdj)*50
 jhdj=jhdj+1
 jhjy=jhjy+jhadd
 conn.Execute ("update 用戶 set allvalue="&jhjy&",等級="&jhdj&" where 姓名='"&zhbg_name&"'")
 Response.Write "<script language=Javascript>{alert('對["&zhbg_name&"]使用了升級卡！國度經驗值上升了"&jhadd&"點，戰鬥等級上升1級，現為"&jhdj&"級');}</script>"
 kapian=zhbg_name&"使用了升級卡，"&zhbg_name&"的泡點增加"&jhadd&"點，戰鬥等級上升1級...真是好福氣呀，不泡點也升級！"
case "健身卡"
 conn.Execute ("update 用戶 set 武功加=武功加+500,內力加=內力加+500,體力加=體力加+500 where 姓名='"&zhbg_name&"'")
 Response.Write "<script language=Javascript>{alert('對["&zhbg_name&"]使用了健身卡，鍛練成功！武功、內力、體力上限值各漲500點！！！');}</script>"
 kapian="【卡片】"&zhbg_name&"大叫著，我要更強，我要更強....從口袋中拿出健身卡，在"&Application("zhbg_user")&"的精心調教下，"&zhbg_name&"經過一翻艱苦訓練，武功、內力、體力上限值各漲500點！！！"
case "好人卡"
	conn.Execute ("update 用戶 set 殺人數=0  where 姓名='" & zhbg_name &"'")
	Response.Write "<script language=JavaScript>{alert('您使用了好人卡，殺人數清0完成！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##說，我是好人，我今天沒有殺人呀……</font>"
case "天眼卡"
	Response.Write "<script>parent.slbox=true;<"&"/script>"	
	Response.Write "<script language=JavaScript>{alert('您使用了天眼卡，現在可以看到別人私聊了！');}</script>"
	kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷地天啟了天眼，十萬百千里的地方都能看見了……</font>"
case "降級卡"
    wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
     rs.close
     rs.open "select * from 用戶 where 姓名='"&to1&"'",conn,2,2
     jhdj=rs("等級")
     del1=jhdj*jhdj*50
     del2=(jhdj+1)*(jhdj+1)*50
     delok=del2-del1
     jhjf=rs("allvalue")-delok
 	conn.Execute ("update 用戶 set 等級=等級-1,allvalue="&jhjf&" where 姓名='" & to1 &"'")
	Response.Write "<script language=JavaScript>{alert('["&to1&"]的等級降為"&jhdj-1&"級了！');}</script>"
	kapian="<img src=card/jiangji.jpg><font color=green><bgsound src=wav/003.mid loop=1><font color=green>【卡片】<font color=" & saycolor & ">##對%%使用了降級卡,%%的等級降為"&jhdj-1&"級，積分也大為減少.</font>"
case "嫁禍卡"
rs.close
rs.open "select * from 用戶 where 姓名='"&to1&"'",conn,2,2
if rs("通緝")="True" then
kapian="<bgsound src=plus/wav/wav/003.mid loop=1><font color=green>【卡片】</font><font color=blue>"&zhbg_name&"</font>對<font color=green>"&to1&"</font>使用<font color=red><b>嫁禍卡</b></font>失敗："&to1&"也是通緝犯，已經不用你嫁禍了！！"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
if rs("門派")="官府" then 
  kapian=""&zhbg_name&"你不可以對管理員使用嫁禍卡……"
  rs.close
  set rs=nothing
  conn.close
  set conn=nothing
  exit function
end if
rs.close
rs.open "select 通緝 from 用戶 where 姓名='"&zhbg_name&"'",conn,2,2
if rs("通緝")="False" then
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>【卡片】</font><font color=blue>"&zhbg_name&"</font>對<font color=green>"&to1&"</font>使用<font color=red><b>嫁禍卡</b></font>失敗："&zhbg_name&"不是通緝犯，怎麼嫁禍呀！！"
        rs.close
set rs=nothing
conn.close
set conn=nothing
exit function
end if
wnky=wnk(to1)
if wnky<>"ok" then 
kapian="<font color=green>【卡片】<font color=" & saycolor & ">##偷偷對%%使用了["&fn1&"]...</font>"&wnky
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
exit function
end if
conn.execute "update 用戶 set 通緝=True where 姓名='"&to1&"'"
conn.execute "update 用戶 set 通緝=False where 姓名='"&zhbg_name&"'"
kapian="<bgsound src=wav/wav/003.mid loop=1><font color=green>【卡片】</font>##偷偷使用卡片：現在%%是通緝犯大家快殺呀！"
Case "豬崽卡"
    rs.close
    rs.Open "select pig from 用戶 where  姓名='" & zhbg_name & "'", conn,2,2
    myhua = rs("pig")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你並沒有小豬,不能使用豬崽卡，請去建設豬圈!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的豬崽數據有問題，請重新建設豬圈!"
                Exit Function
        End If
        '增加豬崽
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 5 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update 用戶 set pig='" & newmyhua & "' where 姓名='" & zhbg_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了豬崽卡<img src=card/zhuzai.JPG><bgsound src=mav/XL.MID loop=1>,得到了5頭豬崽，快去養豬吧！</font>"
Case "種子卡"
    rs.close
    rs.Open "select hua from 用戶 where  姓名='" & zhbg_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" or IsNull(myhua) Or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你並沒有鮮花,不能使用種子卡，請去開墾!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的鮮花數據有問題，請重新開墾!"
                Exit Function
        End If
        '增加種子
        newmyhua = ""
        zt = Split(myhua, "|")
        newmyhua = CLng(zt(0)) + 5 & "|"
        For I = 1 To 26
                newmyhua = newmyhua & zt(I) & "|"
        Next
        conn.Execute "update 用戶 set hua='" & newmyhua & "' where 姓名='" & zhbg_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了種子卡,得到了5顆花種，快去種花吧！</font>"
Case "種花卡"
    rs.close
    rs.Open "select hua from 用戶 where  姓名='" & zhbg_name & "'", conn,2,2
    myhua = rs("hua")
    If myhua = "" Or IsNull(myhua) or InStr(myhua, "|") = 0 Then
                kapian = "<font color=green>【卡片】</font>##你並沒有鮮花,不能使用種花卡，請去開墾!"
                Exit Function
    End If
        zt = Split(myhua, "|")
        ub = UBound(zt)
        If ub <> 27 Then
                kapian = "<font color=green>【卡片】</font>##你的鮮花數據有問題，請重新開墾!"
                Exit Function
        End If
        newmyhua = ""
        For I = 0 To 26
                If I >= 2 Then
                temp = Split(zt(I), ";")
                If CLng(temp(0)) > 0 And CLng(temp(0)) < 94 Then
                ss = CLng(temp(0)) + 5 & ";" & temp(1) & ";0"
                newmyhua = newmyhua & ss & "|"
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
                Else
                newmyhua = newmyhua & zt(I) & "|"
                End If
        Next
        conn.Execute "update [用戶] set hua='" & newmyhua & "' where 姓名='" & zhbg_name & "'"
        kapian="<font color=green>【卡片】<font color=" & saycolor & ">##使用了種花卡,真是爽啊，眼看著花迅速地上漲！</font>"
case else
	Response.Write "<script language=JavaScript>{alert('系統並沒有["&fn1&"]這種卡片,或不能使用你搞錯了！');}</script>"
	Response.End
end select
'刪除自己卡片，記錄
conn.execute "update 用戶 set w5='"&mycard&"' where  姓名='"&zhbg_name&"'"
conn.execute "insert into l(a,b,c,d,e) values (now(),'"& zhbg_name &"','"& to1 &"','操作','"& fn1 & "')"
set rs=nothing	
conn.close
set conn=nothing
end function
'免稅卡
function mhk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from 用戶 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免稅卡")=0 then
		rs.close
	    mhk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免稅卡",1)
		conn.execute "update 用戶 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   'mhk="<br><font color=green>【免稅卡】</font>"&to1&"身上的免稅卡生效,因此不能抓他!"
	   mhk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"正準備喜滋滋的從"&to1&"的口袋中拿錢,就在此時,"&to1&"掏出身上的免稅卡說,慢著,我身上有免稅卡,嘿嘿!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'免罪卡
function mzk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from 用戶 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免罪卡")=0 then
		rs.close
	    mzk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免罪卡",1)
		conn.execute "update 用戶 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   'mzk="<br><font color=green>【免罪卡】</font>"&to1&"身上的免罪卡生效,因此不能抓他!"
	   mzk="<img src=card/myk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&"官府的人準備用根鐵索套在"&to1&"的脖子上,就在此時,"&to1&"掏出身上的免罪卡說,慢著,我身上有免罪卡,嘿嘿!</font>"
	end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'清醒卡
function qxk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from 用戶 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"清醒卡")=0 then
		rs.close
	    qxk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"清醒卡",1)
		conn.execute "update 用戶 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   qxk="><img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"拿出水晶球，睡吧，睡吧，在催眠……"&to1&"睜著個大眼睛，傻嘻嘻的看著他……在說我嗎？"&to1&"掏出身上的清醒卡,慢著,我身上有清醒卡,嘿嘿!"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'免踢卡
function mtk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from 用戶 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"免踢卡")=0 then
		rs.close
	    mtk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"免踢卡",1)
		conn.execute "update 用戶 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   mtk="<img src=card/mhk.jpg></td><td><font color=green><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&zhbg_name&"使出國產臭腳，準備來個國際行動，卻不小心踢到石頭……"&to1&"在一邊嘿嘿的笑，就你呀，還要踢我，再來20年……"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
'萬能卡
function wnk(to1)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("zhbg_usermdb")
  rs.open "select w5 from 用戶 where 姓名='"&to1&"'",conn,2,2
	if iswp(rs("w5"),"萬能卡")=0 then
		rs.close
	    wnk="ok"
	    exit function
	else
		tocard=abate(rs("w5"),"萬能卡",1)
		conn.execute "update 用戶 set w5='"&tocard&"' where  姓名='"&to1&"'"
	   wnk="<img src=card/wangneng.jpg></td><td><font color=green style='font-size=9pt'><bgsound src=wav/003.mid loop=1>【卡片】<font color=" & saycolor & ">"&to1&"在一邊嘿嘿的笑，就你呀，萬能卡，萬能卡，一卡在手，走遍天下……"
	  end if
	  rs.close
	  conn.close
	  set rs=nothing
	  set conn=nothing
end function
%>