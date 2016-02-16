<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
yn=LCase(trim(request.querystring("yn")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
too=trim(Application("ljjh_hunyin"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名 FROM 用戶 WHERE id="&id,conn
peiou=rs("姓名")
'if info(0)<>trim(Application("ljjh_hunyin")) then
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script language=JavaScript>{alert('你想作什麼，人家也沒有說取你！');}</script>"
'	Response.End
'end if
if yn=1 then
if info(0)=too then
	Response.Write "<script language=JavaScript>{alert('我答應你的要求了！');}</script>"
	conn.execute "update 用戶 set 配偶='"& peiou &"',結婚時間=date() where id="&info(9)
	conn.execute "update 用戶 set 配偶='"& info(0) &"',結婚時間=date() where id="&id
	hunyin="恭喜：["& info(0) &"]與{"& peiou &"}喜結良緣，大家表示祝賀！！<img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>喜喜</font></marquee>"
   else
  randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚，一定行的   "
		case 2 
			hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚,你是瘋子他是傻,二半弔子少半啦,沒有你就不瘋來,少了他也就不在傻,你在哪來他在哪,你要死了就少倆."
		case 3
			hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚,秋風清，秋月明。落葉聚還散，寒鴉棲復驚。相思相見知何日，此時此夜難為情。這種男人已經難找了，你就嫁給他吧!"
		case 4
			hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚,我太佩服你了，而且敢這麼說的一定是勇敢的人。我平時太靦腆，一見到異性就臉紅，真沒辦法。"
		case 5
			hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚,等到一切都看透，希望她能陪你看細水常流！如果她不答應你,那可能是她想和我看。"
		case else
			hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚，天荒地老，希望你們情不變 "
	end select
	'if peiou=info(0) then
	'		hunyin=info(0)&"說：支持："&info(0)&"我支持["&peiou&"與"&too&"]結婚，真是不要臉，哪裡有自己說自己好的！"
	'end if
	end if

else
if info(0)=too then
	Response.Write "<script language=JavaScript>{alert('我才不干呢！');}</script>"
	conn.execute "update 用戶 set 魅力=魅力-100 where id="&id
	randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				hunyin="【拒婚】"&info(0)&"對"&peiou&"說：你可以找到比我更好的人  我不想再欺騙你和自己，希望你原諒我吧！我會永遠祝福你的!["& peiou &"]向{"& info(0) &"}求婚不成，魅力下將100點!"
			case 2
				hunyin="【拒婚】"&info(0)&"對"&peiou&"說：我家裡的狗不太喜歡你，你家裡的貓不太喜歡我，所以不要再說什麼，以後你慢慢就明白的!"
			case 3
				hunyin="【拒婚】"&info(0)&"對"&peiou&"說：不知何時，愛情悄悄褪色，失去已無法挽回.....做朋友好嗎？"
			case 4
				hunyin="【拒婚】"&info(0)&"對"&peiou&"說：曾經的滄海已經退潮，我的心裡已經不再有激情，請你還是找別的人吧。"
			case else
				hunyin="【拒婚】"&info(0)&"對"&peiou&"說：你看看我的樣，你也配，我就是嫁給豬無能，也不會嫁給你這個鱉三 "
			end select
'hunyin="【拒婚】"&info(0)&"對"&peiou&"說：你可以找到比我更好的人  我不想再欺騙你和自己，希望你原諒我吧！我會永遠祝福你的!["& peiou &"]向{"& info(0) &"}求婚不成，魅力下將100點!"
else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			hunyin=info(0)&"說：反對："&info(0)&"我反對["&peiou&"與"&too&"]結婚，他們根本不合適   還不如我！"
		case 2
			hunyin=info(0)&"說：反對："&info(0)&"我反對["&peiou&"與"&too&"]結婚，別嫁給他，你看我比他英俊多了，我有車子，房子，票子，現在就缺妻子了。然後ZZZ撇撇嘴想：呵呵~~~先騙騙你再說~~~"
		case 3
			hunyin=info(0)&"說：反對："&info(0)&"我反對["&peiou&"與"&too&"]結婚，他是色狼！前幾天我還看到他從麗春院出來  不行！！"
		case 4
			hunyin=info(0)&"說：反對："&info(0)&"我反對["&peiou&"與"&too&"]結婚，他是色狼！前幾天我還看到他從麗春院出來  不行！！"
		case 5
			hunyin=info(0)&"說：反對："&info(0)&"我反對["&peiou&"與"&too&"]結婚，他虛情假意 恐怕隻會拍馬屁 !"
		case else
			hunyin=info(0)&"說：反對："&info(0)&"我反對["&peiou&"與"&too&"]結婚，嫁給他前認為美麗的事，嫁給他後便完全變為丑惡；山盟海誓的真情，也變成了虛情假意的欺騙，所以請你千萬不要嫁給他!"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【江湖消息】</font>"&hunyin
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
'Application("ljjh_hunyin")=""
Application.UnLock
%>
 