<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
id=Trim(Request.QueryString("id"))
%><html>
<head>
<title>出錯提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="chat/readonly/style.css">

</head>
<body background=linjianww/0064.jpg  leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#dcd8d0">
<tr>
<td background="2.jpg">
<table border="0" cellspacing="0" cellpadding="2" width="346" background="ff.jpg">
<tr>
<td width="324"><font color="FFFFFF"> 出錯提示</font></td>
<td width="14">
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr>
<td width="16"><b><a href="javascript:window.close()" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="關閉"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="343" cellpadding="4" background="1.jpg">
<tr>
<td width="76" align="center" valign="top"><img src="Error.gif" width="76" height="76"></td>
<td width="245">
<%Select Case id
Case "000"%>
<p>  對不起，程序所在目錄不是虛擬目錄，或設置有錯誤，Global.asa 沒有被執行。本程序需要虛擬目錄的支持！如果非虛擬目錄請將global.asa復制到根目錄上去！</p>
<%Case "002"%>
<p>  對不起，本程序不能在這臺服務器上運行，如果你是本程序的合法用戶，請立即與作者聯繫。<br>
信箱：<a href="mailto:seven.s-888@yahoo.com.tw">seven.s-888@yahoo.com.tw</a><br>
</p>
<%Case "003"%>
<p>  數據庫尚未打開，可能是站長正在整理壓縮數據庫，請稍候再來。</p>
<%Case "010"%>
<p>  對不起，本程序專門針對 Internet Explorer 4.0 以上版本的瀏覽器制作，請使用 IE 瀏覽器來訪問。</p>
<%Case "011"%>
<p>  對不起，請勿通過代理服務器來訪問本站。</p>
<%Case "50"%>
<p>  您的oicq格式不正確，oicq為：5-8位的數字！</p>
<%Case "51"%>
<p>  您的Email格式不正確，Email為找回密碼用，非常重要！</p>
<%Case "52"%>
<p>  為了您的安全，密碼請使用6位以上的字符，可以用字母、數字。</p>
<%Case "53"%>
<p>  請用中文輸入數據！</p>
<%Case "54"%>
<p>  系統禁止了“'”,“,”,“"”,“or”等對系統有破壞性的字符！</p>
<%Case "55"%>
<p>  系統為了防止黑客使用探測軟件所以請60秒後再測試新用戶名！</p>
<%Case "56"%>
<p>  用戶名，oicq號，Email，密碼，提示都不能為空！</p>
<%Case "57"%>
<p>  介紹人的姓名不能與你自己的姓名相同！</p>
<%Case "58"%>
<p>  介紹人並不存在，輸入錯誤！</p>
<%Case "59"%>
<p>  注冊人與介紹人為同一ip,系統禁止注冊！</p>
<%Case "60"%>
<p>  介紹人含有非法字符！隻能使用中文（不能帶空格及其它符號）。</p>
<%Case "61"%>
<p>  你不能操作，請退出江湖聊天室再進行！</p>
<%Case "62"%>
<p>  新用戶名已經存在！</p>
<%Case "63"%>
<p>  原用戶名已經不存在或密碼不正確！</p>
<%Case "64"%>
<p>  你是官府中人或者你是門派掌門，禁止改名！</p>
<%Case "65"%>
<p>  你是江湖會員，為了方便管理理，禁止改名！</p>
<%Case "66"%>
<p>  提示與答案不能相同！</p>
<%Case "67"%>
<p>  年為4位數，如2001，月為2位數如，11，日為2位數，如01.</p>
<%Case "68"%>
<p>  所輸入id並不存在！</p>
<%Case "69"%>
<p>  所輸入生日不正確，我們無法查找！</p>
<%Case "70"%>
<p>  所輸入答案不正確，不能取回！</p>
<%Case "71"%>
<p>  你不能操作，請退出江湖聊天室再進行！</p>
<%Case "72"%>
<p>  你想作什麼？呵。。黑老大！</p>

<%Case "100"%>
<p>  歡迎您的光臨，隻是站長已經關閉了聊天室的登錄功能，現在不能進行登錄，請稍後再來。</p>
<%Case "101"%>
<p>  歡迎您的光臨，聊天室現在已滿，站長限制最多同時在線人數為 <font color=red>300</font> 人，請稍後再來。</p>
<%Case "102"%>
<p>  站長禁止新用戶名登錄，請稍後再來。</p>
<%Case "110"%>
<p>  您現在的 IP：<font color=red><%=Request.ServerVariables("REMOTE_ADDR")%></font> 被封鎖 50000分鐘，不能進入聊天室。<br>  離自動解鎖時間還有：<font color=red><%=ABS(50000-int(Datediff("s",Request.QueryString("lockdate"),now())/60))%></font> 分鐘。</p>
<%Case "111"%>
<p>  您現在的 IP：<font color=red><%=Request.ServerVariables("REMOTE_ADDR")%></font> 被<font color=red>【永久】</font>封鎖！不能進入聊天室。請與站長聯繫。</p>
<%Case "120"%>
<p>  用戶名含有非法字符！隻能使用中文（不能帶空格及其它符號,請改名）。</p>
<%Case "121"%>
<p>  密碼含有非法字符！隻能使用英文字母和數字（不能帶空格）。</p>
<%Case "122"%>
<p>  稱謂含有非法字符！隻能使用中文、英文字母和數字（不能帶空格）。</p>
<%Case "123"%>
<p>  新密碼含有非法字符！隻能使用英文字母和數字（不能帶空格）。</p>
<%Case "124"%>
<p>  確認密碼含有非法字符！隻能使用英文字母和數字（不能帶空格）。</p>
<%Case "125"%>
<p>  用戶名的長度超過 10 個字符（一個漢字占兩個字符）。</p>
<%Case "126"%>
<p>  稱謂的長度超過 4 個字符（一個漢字占兩個字符）。</p>
<%Case "127"%>
<p>  用戶名不能為空。</p>
<%Case "128"%>
<p>  密碼不能為空。</p>
<%Case "129"%>
<p>  密碼不能和用戶名相同。</p>
<%Case "130"%>
<p>  對不起，該用戶名為系統所保留，您不能使用這個名字登錄。</p>
<%Case "131"%>
<p>  對不起，該用戶名含有不雅字眼，您不能使用這個名字登錄。</p>
<%Case "132"%>
<p>  對不起，稱謂含有不雅字眼，您不能使用這個稱謂登錄。</p>
<%Case "140"%>
<p>  對不起，該用戶名正在使用中，不能完成您所需要的操作。<br>  如果您是首次使用該用戶名登錄，則可能是該用戶名已經被其他人搶注了，您隻能換用其他名字登錄。<br>  如果您以前曾經使用過該用戶名登錄成功，則可能是您的用戶名被其他人盜用。<br>
另一種可能是：您沒有正常退出聊天室（如：掉線、超時與服務器斷開連接等），請使用掉線處理！導致用戶名被卡在聊天室中，則請您關閉所有瀏覽器，再重新打開，並等十分鐘後再來登錄。如果實在不行，請到留言簿給版主留言，請版主為您解決。</p>
<%Case "141"%>
<p>  對不起，密碼錯誤（注意：密碼區分大小寫），不能完成您所需要的操作。<br>  如果您是首次使用該用戶名登錄，則可能是該用戶名已經被其他人搶注了，您隻能換用其他名字登錄。<br>  如果您以前曾經使用過該用戶名登錄成功，則可能是您的用戶名被其他人盜用，並且盜用者更改了密碼。如果是這種情況，請到留言簿給版主留言，請版主為您解決。</p>
<%Case "142"%>
<p>  對不起，該用戶名被禁用！請到留言簿給版主留言，請版主為您解決。</p>
<%Case "143"%>
<p>  對不起，該用戶名被踢出聊天室，5 分鐘內不能使用該用戶名登錄。還有 <font color=red><%=ABS(300-Datediff("s",Request.QueryString("lastkick"),now()))%></font> 秒。</p>
<%Case "150"%>
<p>  對不起，該用戶名尚未被注冊，當然不能修改密碼了。</p>
<%Case "151"%>
<p>  對不起，您沒有指定“新密碼”，怎麼修改密碼呢？</p>
<%Case "152"%>
<p>  “新密碼”與舊密碼完全相同，用不著再修改密碼啦！</p>
<%Case "153"%>
<p>  “新密碼”不能與用戶名相同！</p>
<%Case "160"%>
<p>  對不起，該用戶名尚未被注冊，當然不能“自殺”了。</p>
<%Case "161"%>
<p>  “確認密碼”為空，不能執行自殺操作。</p>
<%Case "162"%>
<p>  “確認密碼”與“密碼”不一致，不能執行自殺操作。</p>
<%Case "163"%>
<p>  該用戶名被永久保留，不能執行自殺操作。</p>
<%Case "164"%>
<p>  輸入有誤，不能執行自殺操作。</p>
<%Case "165"%>
<p>  有沒有搞錯，你想搞謀殺啊！</p>
<%Case "166"%>
<p>  兩次密碼必需相同才能注冊！</p>
<%Case "167"%>
<p>  用戶名己經被注冊，請選用其他用戶名！</p>

<%Case "001"%>
<p>  該程序執行了非法操作，請立即停止使用，否則可能引起不可預測的後果，請立即與作者取得聯繫。<br>
信箱：<a href="mailto:seven.s-888@yahoo.com.tw?subject=有關笑傲江湖的問題">seven.s-888@yahoo.com.tw</a><br>
</p>
<%Case "200"%>
<p>  沒有此類可供顯示的留言數據，不能查看。</p>
<%Case "201"%>
<p>  沒有此類可供顯示的數據，不能查看。</p>
<%Case "210"%>
<p>  你尚未登錄，或已經超時斷開了連接，請重新登錄。</p>
<%Case "220"%>
<p>  “來自何方”含有非法字符，也不能使用HTML語法！</p>
<%Case "221"%>
<p>  “E-Mail”地址含有非法字符，也不能使用HTML語法！</p>
<%Case "222"%>
<p>  “主頁地址”含有非法字符，也不能使用HTML語法！</p>
<%Case "223"%>
<p>  “來自何方”長度超過30字符，1個漢字占2字符。</p>
<%Case "224"%>
<p>  “E-mail”長度超過50字符。</p>
<%Case "225"%>
<p>  “主頁地址”長度超過50字符 。</p>
<%Case "226"%>
<p>  “個人簡介”長度超過200字符 。</p>
<%Case "227"%>
<p>  “E-mail”地址格式有誤。</p>
<%Case "228"%>
<p>  “主頁地址”格式有誤 。</p>
<%Case "229"%>
<p>  用戶名不存在，不能修改個人信息。</p>
<%Case "230"%>
<p>  請輸入要查詢的用戶名。</p>
<%Case "231"%>
<p>  用戶名：<font color=red><%=Request.QueryString("name")%></font> 不存在。</p>
<%Case "240"%>
<p>  關鍵詞為空，不能搜索。</p>
<%Case "250"%>
<p>  所有“紅色”的項目均為必填項目，請填寫完整。</p>
<%Case "251"%>
<p>  用戶名：<font color=red><%=Request.QueryString("name")%></font> 不存在，不能使用“悄悄話”方式。</p>
<%Case "252"%>
<p>  密碼為空，不能使用用戶名：<font color=red><%=Request.QueryString("name")%></font> 進行留言。</p>
<%Case "253"%>
<p>  密碼錯誤，不能使用用戶名：<font color=red><%=Request.QueryString("name")%></font> 進行留言。</p>
<%Case "254"%>
<p>  “寫給”的長度超過20個字符（1個漢字=2個字符）。</p>
<%Case "255"%>
<p>  “主題”的長度超過40個字符（1個漢字=2個字符）。</p>
<%Case "256"%>
<p>  “內容”的長度超過1024個字。</p>
<%Case "257"%>
<p>  “姓名”的長度超過20個字符（1個漢字=2個字符）。</p>
<%Case "258"%>
<p>  “地址”的長度超過20個字符（1個漢字=2個字符）。</p>
<%Case "259"%>
<p>  “信箱”的長度超過50個字符（1個漢字=2個字符）。</p>
<%Case "260"%>
<p>  “主頁名稱”的長度超過24個字符（1個漢字=2個字符）。</p>
<%Case "261"%>
<p>  “主頁地址”的長度超過50個字符。</p>
<%Case "262"%>
<p>  “信箱”輸入有誤，請重新輸入。</p>
<%Case "263"%>
<p>  “主頁地址”的格式有誤。</p>
<%Case "264"%>
<p>  留言“內容”中不能出現連續的空白行。</p>
<%Case "265"%>
<p>  不能重復粘貼相同的留言。</p>
<%Case "266"%>
<p>  “主題”不能包含半角的“"”“'”引號。</p>
<%Case "300"%>
<p>  該用戶名已經在聊天室中，不能重復進入,請使用<a href=yamen/close.asp>掉線管理</a>再進入江湖！。</p>
<%Case "301"%>
<p>  不能以“江湖管理員”的名稱進入聊天室中，請重回登錄頁面換名登錄。</p>
<%Case "400"%>
<p>  對像：<font color=red><%=Request.QueryString("name")%></font> 不在聊天室中，不能發送消息。</p>
<%Case "401"%>
<p>  消息的長度必須小於 1024 個字符。</p>
<%Case "402"%>
<p>  消息中不能出現連續的空白行。</p>
<%Case "403"%>
<p>  消息不能為空。</p>
<%Case "410"%>
<p>  對像：<font color=red><%=Request.QueryString("name")%></font> 不在聊天室中，不能為其點歌。</p>
<%Case "420"%>
<p>  由於你做壞事被官府人員或系統自動抓進牢裡，等著別人來保釋你吧！不要再干壞事了！</p>
<%Case "421"%>
<p>  你被江湖中人<font color=CCFFCC>[<%=Request.QueryString("xiongshou")%>]</font>打死了，還是到<a href="yamen/disp.asp" target="_blank">閻王殿</a>來吧！</p>
<%Case "422"%>
<p>  由於你被抓進牢裡，等10分鐘後被釋放吧，不要再干壞事了！</p>
<%Case "423"%>
<p>  由於你未注冊、者賬號被盜改名首或者沒有及時還貸款被殺！！請重新<a href='yamen/joinjh.asp'>注冊</a>你的賬號吧！</p>
<%Case "424"%>
<p>  由於你被人點穴！一時間還沒有醒來過了！</p>
<%Case "425"%>
<p>  你不是官府的人或者管理級不夠，你沒有權利訪問這裡！</p>
<%Case "426"%>
<p>  我不是長老，你沒有權利訪問問這裡！</p>
<%Case "427"%>
<p>  你不是江湖中人，不能結婚！</p>
<%Case "428"%>
<p>  怎麼回事？不會是同性戀吧！</p>
<%Case "429"%>
<p>  你來慢一步了  </p>
<%Case "430"%>
<p>  你已經有伴侶了嘛，還想結婚？</p>
<%Case "431"%>
<p>  不能登記！請檢查你的姓名和密碼,是否與登陸時相同！</p>
<%Case "432"%>
<p>  你的密碼不能為空！</p>
<%Case "433"%>
<p>  伴侶的名字不能為空！</p>
<%Case "4333"%>
<p>  你的等級不夠，你想作什麼，申請配偶需要3級以上！</p>
<%Case "4334"%>
<p>  她(他)的等級不夠，你想作什麼，申請配偶需要3級以上！</p>
<%Case "4335"%>
<p>  江湖上找不到你想取的人！</p>
<%Case "4336"%>
<p>  call!你想作什麼？這裡不歡迎你，滾！</p>
<%Case "434"%>
<p>  你想自己取自己！！古往今來都沒有這回事！</p>
<%Case "435"%>
<p>  開什麼玩笑，你們倆性別一樣啊！江湖裡是不允許同性戀的。</p>
<%Case "436"%>
<p>  隻聽裡面傳來一片尖叫，你慌慌張張的逃了出來。</p>
<%Case "437"%>
<p>  你好像沒有帶那麼多錢了吧？</p>
<%Case "438"%>
<p>  對不起，你今天已經很干淨了，溫泉浴每天隻可以洗一次。</p>
<%Case "439"%>
<p>  你沒有登錄或你不是官府中人！你不能進入管理區。</p>
<%Case "440"%>
<%if session("nowinroom")<>"" then
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
ubb=UBound(onlinelist)
for i=1 to ubb step 6
if CStr(onlinelist(i+1))<>CStr(Session("ljjh_name")) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
Session.Abandon

%>
<p>SORRY！出錯以下2種可能性錯誤：<br>1、您沒有注冊登錄！<br>2、您在江湖中被殺！<br>請重新<a href=index.asp>登錄</a>或</a>或<a href=yamen/disp.asp>復活！</a></p>
<%else%>
<p>SORRY！出錯以下2種可能性錯誤：<br>1、您沒有注冊登錄！<br>2、您在江湖中被殺！<br>請重新<a href=index.asp>登錄</a>或<a href=yamen/disp.asp>復活！</a></p>
<%end if%>
<%Case "441"%>
<p>  未找到掌門的資料，掌門未定！</p>
<%Case "442"%>
<p>  錯誤：以上項目中門派、適合必須填寫，適合必須為男、女、所有！</p>
<%Case "443"%>
<p>  錯誤：觀看概率超出範圍，長老，看好再填啊！</p>
<%Case "444"%>
<p>  錯誤：幫派名稱已經存在！</p>
<%Case "445"%>
<p>  錯誤：幫派名稱必須填寫！！</p>
<%Case "446"%>
<p>  未找到掌門的資料，掌門已經被打死了？？？.....不會吧，這麼菜的掌門！</p>
<%Case "447"%>
<p>  該門派已不存在！請刷新頁面！</p>
<%Case "448"%>
<p>  官府不允許改名換姓！</p>
<%Case "449"%>
<p>  錯誤：你不是江湖人,請不要誤闖禁區！</p>
<%Case "449"%>
<p>  未找到掌門的資料，掌門未定！</p>
<%Case "450"%>
<p>  你幫今日己領了一次銀兩了，金庫不再支付你的開銷，省著點吧！</p>
<%Case "451"%>
<p>  錯誤：你不是掌門,請不要誤闖禁區！</p>
<%Case "452"%>
<p>  錯誤：你是官府的人，不可以改名！</p>
<%Case "453"%>
<p>  錯誤：你不是江湖中人，請不要隨便亂闖！</p>
<%Case "454"%>
<p>  錯誤：你不是江湖中人，我們不收你這種打短工的小人物！</p>
<%Case "455"%>
<p>  錯誤：你無門無派，不許進入本幫禁地！</p>
<%Case "456"%>
<p>  錯誤：這個漏洞己堵上，請不要再試了！</p>
<%Case "457"%>
<p>  錯誤：沒錢也想改名？？先賺點銀子再來改名吧！</p>
<%Case "458"%>
<p>  錯誤：你的銀兩不夠，我們不接受你的離婚登記！</p>
<%Case "459"%>
<p>  錯誤：操作不成功!!請返回！</p>
<%Case "460"%>
<p>  錯誤：你等級太低了點吧！要是創派後弟子都養不活吧！</p>
<%Case "461"%>
<p>  錯誤：你好像有門派了吧！有多大能奈，還想建多少個派呀！</p>
<%Case "462"%>
<p>  錯誤：你的銀兩買不起這件物品！</p>
<%Case "463"%>
<p>  錯誤：你沒有寄售的物品在本集市！</p>
<%Case "464"%>
<p>  錯誤：輸入的內容出錯，請確認您輸入的號碼為數字！</p>
<%Case "465"%>
<p>  對不起，由於您行為不端，道德太低，罰你到<a href="xg.asp">思過壁</a>“面壁思過”。</p>
<%Case "466"%>
<p>  對不起，因錢莊小本經營，您的現金數目太多，請下次再來！</p>
<%Case "467"%>
<p>  江湖高手，你武功超過150000上限，系統已經自動將其降至1499999，請不要再提升武功了，你已經是江湖一等一高手了，請<a href="http://xajh.us">重新登錄</a>即可！</p>
<%Case "468"%>
<p>  江湖高手，你體力超過1億上限，系統已經自動將其降至99000000，請不要再提升體力了，你已經是江湖一等一高手了，請<a href="http://xajh.us">重新登錄</a>即可！</p>
<%Case "469"%>
<p>  你沒有這種類型的裝備，請到集市購買相關裝備</p>
<%Case "470"%>
<p>  你沒有任何物品。</p>
<%Case "471"%>
<p>  不能設置負數參數。</p>
<%Case "472"%>
<p>  由於你在休息中，現在還沒有醒來或被<font color=CCFFCC>[<%=Request.QueryString("xiongshou")%>]</font>使用了水晶球，過2分鐘就會醒來！</p>
<%Case "473"%>
<p>  由於你違反江湖規則，已被終生監禁，此賬號不能使用了！</p>
<%Case "474"%>
<p>  由於你人品太差，道德小於-10000，江湖不歡迎你，請去采冰、采礦加錢去孤兒院買道德去吧！！</p>
<%Case "475"%>
<p>  對不起，你輸入的數據有一項是空值！</p>
<%Case "476"%>
<p>  對不起,官府目前並未出題給你答！</p>
<%Case "477"%>
<p>  對不起,為了公平,官府的人不能參於此活動！</p>
<%Case "478"%>
<p>  呵...黑客老大,這裡不歡迎你過來,想玩去別人家吧!</p>
<%Case "479"%>
<p>  對不起,你所提供的答案不對！你得不到獎勵！</p>
<%Case "480"%>
<p>  對不起,您被點穴了，過幾秒鐘會穴道會自動解開,請一會再登錄！</p>
<%Case "481"%>
<p>  想作夢取媳婦呀？好好想想吧你！</p>
<%Case "482"%>
<p>  你的等級太低了，代不了這些款！</p>
<%Case "483"%>
<p>  你已經代過款了，請不要理貸款了！</p>
<%Case "484"%>
<p>  錯誤，沒有該類商店！</p>
<%Case "485"%>
<p>  錯誤：輸入的號碼不是5位，彩票隻發行5位的數字，請重新輸入！</p>
<%Case "486"%>
<p>  錯誤：輸入的內容出錯，請確認您輸入是5位整數!</p>
<%Case "487"%>
<p>  錯誤：輸入的內容出錯，請確認您輸入的號碼為數字！</p>
<%Case "488"%>
<p>  錯誤：請你填入您想購買的彩票號碼！</p>
<%Case "489"%>
<p>  錯誤：對不起，您的號碼己紀被人買走了，請選用其他的號！</p>
<%Case "490"%>
<p>  錯誤：對不起，您今天己經購買了彩票了！</p>
<%Case "900"%>
<p>  錯誤：對不起，你的管理級不夠操作！</p>
<%Case "1000"%>
<p>  錯誤：您正在使用的靈劍江湖不是正版，請支持正版江湖，靈劍總站：<a href="http://xajh.us">http://xajh.us</a>！</p>
<%Case else%>
<p>  對不起，該出錯類型未被登記，請聯繫站長解決。</p>
<%End Select
%>
</td>
</tr>
<tr>
<td colspan="2" align="center" valign="top">
<input type="button" name="ok" value=" 返 回 " onclick=javascript:history.go(-1);>
<input type="button" name="ok2" value=" 關 閉 " onclick=javascript:window.close();>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>
