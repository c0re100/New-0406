<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
id=request.querystring("id")
%>
<html><head><meta http-equiv='content-type' content='text/html; charset=big5'>
<style type='text/css'>
.webstyle   {font-family: Webdings; font-size: 9pt}
.yy4{ font-size: 9pt;color:FFFFFF; line-height: 11pt;position: relative; width: 100% }
body{font-size:9pt;}input{font-size:9pt;}a{font-size:9pt;color:FFFF00;text-decoration:none;}a:hover{color:FFFF00;text-decoration:underline;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=000000 background=http://xs68.xs.to/pics/06072/b39.gif  bgproperties=fixed topmargin=0 text=FFFF00>
<form name=af method=POST action='say.asp'  target='d' onsubmit='return(parent.checksays());'>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr><td>
<div align=center><input type=hidden name='mdtx' value="<%=info(10)%>">
<input type=hidden name='grade' value="<%=info(2)%>">
<input type=hidden name='fs' value='10'>
<input type=hidden name='lh' value='125'>
<input type=hidden name='sy' value=''>
<input type=hidden name='oldsays' value>
<input type=hidden name='oldact' value>
<input type=hidden name='oldtowho' value>
<select name='addwordcolor' style='font-size:9pt;background-color:#000000'>
<option style="color:00FF00" value="00FF00" selected>名</option>
<option style="color:00FF00" value="00FF00">名</option>
<option style="color:0000FF" value="0000FF">名</option>
<option style="color:FFFF00" value="FFFF00">名</option>
<option style="color:FFCC00" value="FFCC00">名</option>
<option style="color:FFD0CF" value="FFD0CF">名</option>
<option style="color:F0E0A0" value="F0E0A0">名</option>
<option style="color:B0DF90" value="B0DF90">名</option>
<option style="color:90CFEF" value="90CFEF">名</option>
<option style="color:CF9FC0" value="CF9FC0">名</option>
<option style="color:C0C0C0" value="C0C0C0">名</option>
<option style="color:FFFFFF" value="FFFFFF">名</option>
<option style="color:A992D6" value="A992D6">名</option>
<option style="color:53A6A6" value="53A6A6">名</option>
</select>
<select name='sayscolor'  style='font-size:9pt;background-color:#000000'>
<option style="color:FFCC00" value="FFCC00" selected>話</option>
<option style="color:00FF00" value="00FF00">話</option>
<option style="color:0000FF" value="0000FF">話</option>
<option style="color:FFFF00" value="FFFF00">話</option>
<option style="color:FFCC00" value="FFCC00">話</option>
<option style="color:F0E0A0" value="F0E0A0">話</option>
<option style="color:B0DF90" value="B0DF90">話</option>
<option style="color:90CFEF" value="90CFEF">話</option>
<option style="color:CF9FC0" value="CF9FC0">話</option>
<option style="color:C0C0C0" value="C0C0C0">話</option>
<option style="color:FFFFFF" value="FFFFFF">話</option>
<option style="color:A992D6" value="A992D6">話</option>
<option style="color:53A6A6" value="53A6A6">話</option>
</select>
<select name='addsays' onchange="document.af.sytemp.focus();" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>表情</option>
<option value="微笑地">微笑</option>
<option value="溫柔地">溫柔</option>
<option value="紅著臉">臉紅</option>
<option value="搖頭晃腦得意地">得意</option>
<option value="哈！哈！哈！笑著">大笑</option>
<option value="快要哭地">快哭</option>
<option value="哭著">哭</option>
<option value="不懷好意地">壞意</option>
<option value="遺憾地">遺憾</option>
<option value="瞪大了眼睛，很詫異地">詫異</option>
<option value="幸福地">幸福</option>
<option value="悲痛地">悲痛</option>
<option value="正義凜然地">正義</option>
<option value="嚴肅地">嚴肅</option>
<option value="生氣地">生氣</option>
<option value="大聲地">大聲</option>
<option value="傻乎乎地">傻</option>
<option value="很滿足地">滿足</option>
<option value="很無辜地">無辜</option>
<option value="依依不舍地">不舍</option>
<option value="口吐白沫">白沫</option>
<option value="無可奈何地">無奈</option>
<option value="可憐兮兮地">可憐</option>
</select>
<input type=text name='username' readonly  style="border:1px solid #6699FF; text-align:center;font-size:9pt;color:#FFCCCC;background-color:#000000" size=5 maxlength=10>→
<select name='towho' style='font-size:12px;color:#FFFFFF;background-color:000000' onClick=dj()><option value='大家' selected>大家</option></select>
<input type=text name='sytemp'   style='font-size:9pt;color:ffffff;background-color:#000000' size=30 maxlength=180 >
<a href="#" onClick="gop()"><img src=ico/wfy_i_c_g_b32.gif border=0></a>
<a href="#" onClick="gon()"><img src=ico/wfy_i_c_g_b31.gif border=0></a>
<input type=submit  name='subsay' value='發言' style="background-color:FF0000;color:#FFFFFF;border: 1 double" onmouseover="this.style.color='FFFF00'" onmouseout="this.style.color='FFFFFF'">
<input type=text name='clock' style="text-align:right;font-size:9pt;background-color:#000000;color:#FFFFFF;border: 1px double #3399FF; " value="" size=4 readonly>
</td></tr>
</table>
<div align=center>
<input name="tu" type="hidden" value="">
<input name="addsign" type="hidden" value=0>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option selected>心情
<option style=background-color:"#000000" value="/心情$ 正常">正常
<option style=background-color:"#000000" value="/心情$ 泡點">泡點
<option style=background-color:"#000000" value="/心情$ 自定">自定
</select>
<select name='command' onchange="rc(this.value);" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>使用命令1 </option>
<option style="color:#FFFFFF" value="/信息$">查看信息 </option>
<option style="color:#FFFFFF" value="/搶奪$">搶奪寶物 </option>
<option style="color:#FFFFFF" value="/單挑$ 挑戰文書">比武單挑 </option>
<option style="color:#FFFFFF" value="/求婚$ 真愛表白">示愛求婚 </option>
<%if info(1)="男" then %>
<option style="color:#FFFFFF" value="/求妾$">求小老婆 </option>
<% else %>
<option style="color:#FFFFFF" value="/求小老公$">求小老公 </option>
<%end if%>
<option style="color:#FFFFFF" value="/同意$">答應過門 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/主題$ ">千裡傳音 </option>
<option style="color:#FFFFFF" value="/心動$ ">心動感覺 </option>
<option style="color:#FFFFFF" value="/怒吼$ ">狂獅怒吼 </option>
<option style="color:#FFFFFF" value="/心跳$ ">心跳心語 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/基金$">本門基金 </option>
<option style="color:#FFFFFF" value="/修練$ ">寶物修練 </option>
<option style="color:#FFFFFF" value="/暴豆$ ">暴威力豆 </option>
<option style="color:#FFFFFF" value="/拜師$ ">拜師習武 </option>
<option style="color:#FFFFFF" value="/下山$ ">學成下山 </option>
<option style="color:#FFFFFF" value="/收徒$ ">招收徒弟 </option>
<option style="color:#FFFFFF" value="/加入$ 門派名">加入門派</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/傳授$ 100">傳授內力 </option>
<option style="color:#FFFFFF" value="/經驗$ 100">傳授經驗 </option>
<option style="color:#FFFFFF" value="/送錢$ 1000">贈送現金 </option>
<option style="color:#FFFFFF" value="/贈送$ 寫出物品名">贈送物品 </option>
<option style="color:#FFFFFF" value="/丟棄$ 寫出物品名">丟棄物品 </option>
<option style="color:#FFFFFF" value="/使用$ 卡片名">使用卡片 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/醫療$ ">妙手回春</option>
<option style="color:#FFFFFF" value="/美容$ ">美容大法</option>
<option style="color:#FFFFFF" value="/補氣$ ">遊雲靈氣</option>
<option style="color:#FFFFFF" value="/教武$ ">教導武功</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/掃黃">掃黃行動</option>
<option style="color:#FFFFFF" value="/賣了她吧$ ">拐賣行動</option>
<option style="color:#FFFFFF" value="/小偷$ ">偷取物品</option>
<option style="color:#FFFFFF" value="/抓小偷$ ">抓小偷</option>
<option style="color:#FFFFFF" value="/孩子偷竊$ ">孩子偷竊</option>
<option style="color:#FFFFFF" value="/教訓孩子$ ">教訓孩子</option>
</select>
<select name='command2' onChange="rc(this.value);" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>使用命令2 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/打撲克$ 1000|銀兩[或：金幣，法力]">打撲克
<option style="color:#FFFFFF" value="/發牌$ 公證|負[或勝]" style=color:#000000>撲公證
<option style="color:#FFFFFF" value="/打麻將$ 1000|銀兩[或：金幣，法力]">打麻將
<option style="color:#FFFFFF" value="/出牌$ 公證|負[或勝]" style=color:#000000>麻公證
<option style="color:#FFFFFF" value="/幸運$ 幸運數字1-1000">靈劍彩票 </option>
<option style="color:#FFFFFF" value="/雙人賭博$ 請填入數字">兩人對賭 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/發招$ 招式名">發招攻擊
<option style="color:#FFFFFF" value="/吸星大法$ ">吸取內力 </option>
<option style="color:#FFFFFF" value="/偷錢$ ">偷取錢財 </option>
<option style="color:#FFFFFF" value="/下毒$ 毒藥名">偷偷下毒 </option>
<option style="color:#FFFFFF" value="/投擲$ 暗器名">投擲暗器 </option>
<option style="color:#FFFFFF" value="/拼命$ ">同歸於盡 </option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/神思$ ">崖上神思 </option>
<option style="color:#FFFFFF" value="/參禪$ ">山洞參禪 </option>
<option style="color:#FFFFFF" value="/打坐$ ">打坐練功 </option>
<option style="color:#FFFFFF" value="/閉目$ ">閉目養神 </option>
<option style="color:#FFFFFF" value="/練武$ ">武館習武 </option>
<option style="color:#FFFFFF" value="/修習武功$ 武功招式">修習武功 </option>
<option style="color:#FFFFFF" value="/修養$ ">修心養性 </option>
</select>
<select name='command3' onChange="rc(this.value);" style='font-size:12px;color:#FFFFFF;background-color:000000'>
<option value="" selected>使用命令3</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/配制寶石$ ">配制寶石</option>
<option style="color:#FFFFFF" value="/魔力鑽石$ ">魔力鑽石</option>
<option style="color:#FFFFFF" value="/生日蛋糕$ ">生日蛋糕</option>
<option style="color:#FFFFFF" value="/發射子彈$ ">發射子彈</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/存錢$ 0">銀行存錢 </option>
<option style="color:#FFFFFF" value="/取錢$ 0">銀行取錢 </option>
<option style="color:#FFFFFF" value="/轉賬$ 10000">銀行轉賬 </option>
<option style="color:#FFFFFF" value="/傳送法力$ 100">傳送法力</option>
<option style="color:#FFFFFF" value="/存法力$ 100">存放法力</option>
<option style="color:#FFFFFF" value="/取法力$ 100">取回法力</option>
<%if info(6)="掌門" or info(6)="副掌門" then%>
<option style="color:#FFFFFF" value="/掌門令$ 公布的消息">掌 門 令 </option>
<option style="color:#FFFFFF" value="/同盟$ ">幫派結盟</option>
<option style="color:#FFFFFF" value="/解除同盟$ ">解除同盟</option>
<option style="color:#FFFFFF" value="/幫派大戰$ ">幫派大戰</option>
<option style="color:#FFFFFF" value="/冊封$ 身份名">冊封身份 </option>
<option style="color:#FFFFFF" value="/冊封$ 副手">冊封副手 </option>
<option style="color:#FFFFFF" value="/冊封$ 長老">冊封長老 </option>
<option style="color:#FFFFFF" value="/冊封$ 護法">冊封護法 </option>
<option style="color:#FFFFFF" value="/冊封$ 堂主">冊封堂主 </option>
<option style="color:#FFFFFF" value="/冊封$ 查看">冊封查看 </option>
<option style="color:#FFFFFF" value="/冊封$ 弟子">取消冊封 </option>
<option style="color:#FFFFFF" value="/門規處置$ 1000">門規處置 </option>
<option style="color:#FFFFFF" value="/招收$ ">招收弟子</option>
<%end if
if info(6)="長老" then%>
</option>
<option style="color:#FFFFFF" value="/冊封$ 護法">冊封護法 </option>
<option style="color:#FFFFFF" value="/冊封$ 堂主">冊封堂主 </option>
<option style="color:#FFFFFF" value="/冊封$ 查看">冊封查看 </option>
<option style="color:#FFFFFF" value="/冊封$ 弟子">取消冊封 </option>
<option style="color:#FFFFFF" value="/門規處置$ 1000">門規處置 </option>
<option style="color:#FFFFFF" value="/啞穴$ 請寫明原因 ">教訓弟子 </option>
<option style="color:#FFFFFF" value="/獎勵$ 5000 ">獎勵弟子
<option style="color:#FFFFFF" value="/招收$ ">招收弟子
<%end if
if  info(6)="護法" then%>
</option>
<option style="color:#FFFFFF" value="/冊封$ 堂主">冊封堂主 </option>
<option style="color:#FFFFFF" value="/冊封$ 查看">冊封查看 </option>
<option style="color:#FFFFFF" value="/冊封$ 弟子">取消冊封 </option>
<option style="color:#FFFFFF" value="/啞穴$ 請寫明原因 ">教訓弟子 </option>
<option style="color:#FFFFFF" value="/獎勵$ 100 ">獎勵弟子
<option style="color:#FFFFFF" value="/招收$ ">招收弟子
<option style="color:#FFFF66">========
<%end if
if  info(6)="堂主" then%>
</option>
<option style="color:#FFFFFF" value="/冊封$ 查看">冊封查看 </option>
<option style="color:#FFFFFF" value="/冊封$ 弟子">取消冊封 </option>
<option style="color:#FFFFFF" value="/招收$ ">招收弟子
<option style="color:#FFFFFF" value="/獎勵$ 100 ">獎勵弟子
<option style="color:#FFFF66">========
<%end if
if info(2)>=6 then%>
</option>
<option style="color:#FFFF66">======</option>
<option style="color:#FFFFFF" value="/單挑取消$ ">取消單挑 </option>
<option style="color:#FFFFFF" value="/復活$ ">使用復活 </option>
<option style="color:#FFFFFF" value="/啞穴$ 請寫明原因 ">使用啞穴 </option>
<option style="color:#FFFFFF" value="/警告$ 請寫明原因 ">使用警告 </option>
<option style="color:#FFFFFF" value="/逮捕$ 請寫明原因 ">使用逮捕
<% end if
if info(2)>=7 then%>
</option>
<option style="color:#FFFFFF" value="/坐牢$ 請寫明原因 ">使用坐牢 </option>
<option style="color:#FFFFFF" value="/踢人$ 請寫明原因 ">使用踢人
<% end if
if info(2)>=8 then%>
</option>
<option style="color:#FFFFFF" value="/點穴$ 請寫明原因 ">使用點穴 </option>
<option style="color:#FFFFFF" value="/解穴$ ">解穴操作 </option>
<option style="color:#FFFFFF" value="/罰款$ 請寫明原因">使用罰款 </option>
<option style="color:#FFFFFF" value="/動刑$ 請寫明原因">使用動刑 </option>
<% end if
if info(2)>=9 then%>
</option>
<option style="color:#FFFFFF" value="/放大$ ">放大 </option>
<option style="color:#FFFFFF" value="/監禁$ 請寫明原因 ">使用監禁 </option>
<option style="color:#FFFFFF" value="/釋放$ 用戶名">解除監禁 </option>
<option style="color:#FFFFFF" value="/禁打開關$">禁打開關 </option>
<% end if
if info(2)>=10 then%>
</option>
<option style="color:#FFFFFF" value="/炸彈$">喫 炸 彈 </option>
<option style="color:#FFFFFF" value="/官府獎勵$ 10000">使用獎勵 </option>
<option style="color:#FFFFFF" value="/公告$ ">官府公告
<option style="color:#FFFFFF" value="/隱身$">隱身開關</option>
<option style="color:#FFFFFF" value="/查ip$ ">查看IP
<% end if 
if info(2)>=11 then%>
<option style="color:#FFFFFF" value="/跟蹤私聊">跟蹤私聊 </option>
<option style="color:#FFFFFF" value="/取消跟蹤">取消跟蹤 </option>
<option style="color:#FFFFFF" value="/查看物品$">查看物品 </option>
<option style="color:#FFFFFF" value="/分類物品$ 卡片">分類查看
<option style="color:#FFFFFF" value="/封鎖ip$ 請寫明原因">封鎖IP </option>
<option style="color:#FFFFFF" value="/解鎖ip$">解鎖IP </option>
<% end if %>
</option>
</select>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:#000000'>
<option selected>常用命令4
<option style="color:#FFFF66">=發光文字=</option>
<option style="color:#FFFFFF" value="/藍光文字$ ">藍光文字 </option>
<option style="color:#FFFFFF" value="/綠光文字$ ">綠光文字 </option>
<option style="color:#FFFFFF" value="/紫光文字$ ">紫光文字 </option>
<option style="color:#FFFFFF" value="/黃光文字$ ">黃光文字 </option>
<option style="color:#FFFFFF" value="<I>打字 ">斜體字 </option>
<option style="color:#FFFFFF" value="<B>打字 ">粗體字 </option>
<option style="color:#FFFFFF" value="<BIG>打字 ">大字體 </option>
<option style="color:#FFFFFF" value="<CENTER>打字 ">向中對齊 </option>
<option style="color:#FFFFFF" value="<u>打字</u> ">底線字 </option>
<option style="color:#FFFFFF" value="<marquee>打字</marquee> ">走馬燈 </option>
<option style="color:#FFFF66">=====</option>
<option style="color:#FFFFFF" value="/送會費$ 10000">贈送會費 </option>
<option style="color:#FFFFFF" value="/送金幣$ 10000">贈送金幣 </option>
<option style="color:#FFFFFF" value="/尋找流星$ ">尋找流星 </option>
<option style="color:#FFFFFF" value="/尋找龍珠$ ">尋找龍珠 </option>
<option style="color:#FFFFFF" value="/尋找流星淚$ ">尋找流星淚 </option>
<option style="color:#FFFF66">=領取數值=</opton>
<option style="color:#FFFFFF" value="/領取體力$ ">領取體力 </option>
<option style="color:#FFFFFF" value="/領取內力$ ">領取內力</option>
<option style="color:#FFFFFF" value="/領取武功$ ">領取武功 </option>
<option style="color:#FFFFFF" value="/領取攻擊$ ">領取攻擊 </option>
<option style="color:#FFFFFF" value="/領取防御$ ">領取防御 </option>
<option style="color:#FFFFFF" value="/領取魅力$ ">領取魅力 </option>
<option style="color:#FFFFFF" value="/領取道德$ ">領取道德</option>
<option style="color:#FFFFFF" value="/領取法力$ ">領取法力 </option>
<option style="color:#FFFFFF" value="/領取資質$ ">領取資質 </option>
<option style="color:#FFFFFF" value="/領取會員費$ ">領取會費 </option>
<option style="color:#FFFFFF" value="/領取金幣$ ">領取金幣 </option>
</select>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:#000000'>
<option selected>常用命令5
<option style="color:#FFFF66">=發光字=</option>
<option style="color:#FFFFFF" value="/藍光字$ ">藍光字 </option>
<option style="color:#FFFFFF" value="/紅光字$ ">紅光字 </option>
<option style="color:#FFFFFF" value="/綠光字$ ">綠光字 </option>
<option style="color:#FFFFFF" value="/黃光字$ ">黃光字 </option>
<option style="color:#FFFFFF" value="/紫光字$ ">紫光字 </option>
<option style="color:#FFFFFF" value="/底色文字$ ">背光字 </option>
<option style="color:#FFFF66">=新增功能=</option>
<option style="color:#FFFFFF" value="/藍光文字$ 去http://lok332005.somee.com，果江湖有好多功能，江湖抽獎，江湖商店，總覽．註冊果時介紹人：ｘｘｘ，註冊左會有１ｌｖ會員&5000金幣">拉人講d咩 </option>
<option style="color:#FFFFFF" value="/藍光文字$ 領取系統*數值商店v2.0*江湖抽獎*自助加管系統*勁多正卡*特別字體*怪術*怪咒*廣告列*唔會彈出街*唔會無法顯示網頁*等等!... ">本網功能 </option>
<option style="color:#FFFFFF" value="/斬首術$ ">斬首術 </option>
<option style="color:#FFFFFF" value="/耿鬼催眠術$ ">耿鬼催眠術 </option>
<option style="color:#FFFFFF" value="/富迪解眠術$ ">富迪解眠術 </option>
<option style="color:#FFFFFF" value="/強制出關術$ ">強制出關術 </option>
<option style="color:#FFFFFF" value="/地獄減管術$ ">地獄減管術 </option>
<option style="color:#FFFFFF" value="/黑暗減攻術$ ">黑暗減攻術 </option>
<option style="color:#FFFFFF" value="/黑暗減防術$ ">黑暗減防術 </option>
<option style="color:#FFFFFF" value="/黑暗減員術$ ">黑暗減員術 </option>
<option style="color:#FFFFFF" value="/黑暗減金術$ ">黑暗減金術 </option>
<option style="color:#FFFFFF" value="/黑暗減會術$ ">黑暗減會術 </option>
<option style="color:#FFFFFF" value="/地獄減錢術$ ">地獄減錢術 </option>
<option style="color:#FFFFFF" value="/光明加管術$ ">光明加管術 </option>
<option style="color:#FFFFFF" value="/光明加員術$ ">光明加員術 </option>
<option style="color:#FFFFFF" value="/光明加防術$ ">光明加防術 </option>
<option style="color:#FFFFFF" value="/光明加攻術$ ">光明加攻術 </option>
<option style="color:#FFFFFF" value="/光明加金術$ ">光明加金術 </option>
<option style="color:#FFFFFF" value="/光明加會術$ ">光明加會術 </option>
<option style="color:#FFFFFF" value="/光明加錢術$ ">光明加錢術 </option>
</select>
<select name='zt' onChange="rc(this.value);"  style='font-size:12px;color:#FFFFFF;background-color:#000000'>
<option selected>購買數值
<option style="color:#FFFF66">==購買數值===</option>
<option style="color:#FFFFFF" value="/購買會費$ ">購買會費 </option>
<option style="color:#FFFFFF" value="/購買金幣$ ">購買金幣 </option>
<option style="color:#FFFFFF" value="/購買子彈$ ">購買子彈 </option>
<option style="color:#FFFFFF" value="/購買資質$ ">購買資質 </option>
<option style="color:#FFFFFF" value="/購買法力$ ">購買法力 </option>
<option style="color:#FFFF66">==購買能力===</option>
<option style="color:#FFFFFF" value="/購買體力$ ">購買體力 </option>
<option style="color:#FFFFFF" value="/購買武功$ ">購買武功 </option>
<option style="color:#FFFFFF" value="/購買攻擊$ ">購買攻擊 </option>
<option style="color:#FFFFFF" value="/購買防御$ ">購買防御 </option>
<option style="color:#FFFFFF" value="/購買魅力$ ">購買魅力 </option>
<option style="color:#FFFFFF" value="/購買道德$ ">購買道德 </option>
<option style="color:#FFFF66">==購買會員===</option>
<option style="color:#FFFFFF" value="/購買2級會員$ ">購買2級會員 </option>
<option style="color:#FFFFFF" value="/購買3級會員$ ">購買3級會員 </option>
<option style="color:#FFFFFF" value="/購買4級會員$ ">購買4級會員 </option>
<option style="color:#FFFFFF" value="/購買5級會員$ ">購買5級會員 </option>
<option style="color:#FFFFFF" value="/購買6級會員$ ">購買6級會員 </option>
<option style="color:#FFFFFF" value="/購買7級會員$ ">購買7級會員 </option>
<option style="color:#FFFFFF" value="/購買8級會員$ ">購買8級會員 </option>
<option style="color:#FFFFFF" value="/購買9級會員$ ">購買9級會員 </option>
<option style="color:#FFFFFF" value="/購買10級會員$ ">購買10級會員 </option>
<option style="color:#FFFF66">==購買官府===</option>
<option style="color:#FFFFFF" value="/購買6級官府$ ">購買6級官府 </option>
<option style="color:#FFFFFF" value="/購買7級官府$ ">購買7級官府 </option>
<option style="color:#FFFFFF" value="/購買8級官府$ ">購買8級官府 </option>
<option style="color:#FFFFFF" value="/購買9級官府$ ">購買9級官府 </option>
<option style="color:#FFFFFF" value="/購買10級官府$ ">購買10級官府 </option>
<option style="color:#FFFF66">==職業轉換===</option>
<option style="color:#FFFFFF" value="/購買水天師$ ">購買水天師 </option>
<option style="color:#FFFFFF" value="/購買神龍戰士$ ">購買神龍戰士 </option>
<option style="color:#FFFFFF" value="/購買金剛勇士$ ">購買金剛勇士 </option>
</select>
<input type=button value='復活' onClick="window.open('../yamen/disp.asp','casper','scrollbars=yes,resizable=yes,width=400,height=300')" style="background-color:FF6633;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button2">
<input type=button value='貼圖' onClick="window.open('pic.asp','f3')" style="background-color:#000000;color:#FFFFFF;border: 1px double #3399FF; " name="button">
<input type=button name="tbclutch" value="全屏" onClick="javascript:parent.tbclutch();" title="合屏/分屏/分屏上大/分屏下大/垂直切換" onMouseOver="window.status='合屏/分屏/上大/下大/垂直切換。';return true" onMouseOut="window.status='';return true" style="background-color:#000000;color:#FFFFFF;border: 1px double #3399FF; " >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> <td> <div align="center"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr> <td height="20">
<div align="center">
<input type="checkbox" name="addvalues">
<a href="#" onClick='document.af.addvalues.checked=!(document.af.addvalues.checked);document.af.sytemp.focus();'>
<font color="#90CFEF">泡點</font></a>
<input type='checkbox' name='sp' accesskey='v' onClick='parent.f1.speed();document.af.sytemp.focus();' value="ON">
<a href=# onClick='document.af.sp.checked=!(document.af.sp.checked);parent.f1.speed();document.af.sytemp.focus();' onMouseOver="window.status='流暢的顯示對話區文字。';return true" onMouseOut="window.status='';return true" title="“啟用/關閉”流暢滾屏方式|快捷鍵Alt+V">
<font color="#90CFEF">快</font></a><a href="#"><font color="#90CFEF">轉</font></a>
<input onclick='document.af.listname.focus();if(document.af.listname.checked==true){parent.m.location.reload();}'
type=checkbox name=listname title='有人進入或退出的時候,自動刷新名單;' value="ON" checked>
<a href='#' onclick='document.af.listname.checked=!(document.af.listname.checked);
document.af.listname.focus();if(document.af.listname.checked==true){parent.m.location.reload();}' title='有人進入或退出的時候,自動刷新名單; '>
<font color="#90CFEF">名單</font></a>
<input onClick='document.af.soundtx.focus();' type=checkbox name=soundtx checked title='是否使用聊天室聲音、音樂等！' value="ON">
<a href='#' onClick='document.af.soundtx.checked=!(document.af.soundtx.checked);document.af.soundtx.focus();' title='是否使用聊天室聲音、音樂等！'>
<font color="#90CFEF">音</font></a><a href="#"><font color="#90CFEF">效</font></a>
<input type=checkbox name='towhoway' value='1' accesskey='s' onClick="document.af.sytemp.focus();">
<a href='#' onMouseOver="window.status='選中本項後，說的話只有你和對方才能看到（即使站長也看不到）。';return true" onMouseOut="window.status='';return true" onClick="document.af.towhoway.checked=!(document.af.towhoway.checked);document.af.sytemp.focus();" title="悄悄話兒悄悄說">
<font color="#90CFEF">私聊</font></a>
<input type='checkbox' name=as accesskey='a' checked onClick='parent.f1.scrollit();document.af.sytemp.focus();' value="ON">
<a href='#' onclick='document.af.as.checked=!(document.af.as.checked);
document.af.as.focus();parent.f1.scrollit();document.af.sytemp.focus();' title='是否滾屏顯示|快捷鍵Alt+A '>
<font color="#90CFEF">滾屏</font></a>
<input type="checkbox" name="titleline" value="1" accesskey='t' onClick="document.af.sytemp.focus();">
<a href='#' onMouseOver="window.status='將發言內容置於聊天室的用戶標題行。';return true" onMouseOut="window.status='';return true" onClick="document.af.titleline.checked=!(document.af.titleline.checked);document.af.sytemp.focus();" title="二級以上聊友才能使用">
<font color="#90CFEF">標題</font></a>
<%if info(2)>=6  then%>
<a href="../dt/ask.asp" target="_blank"><font color="#90CFEF">發題</font></a>
<a href="../tongji/tj.asp" target="_blank"><font color="#90CFEF">通輯</font></a>
<a class=blue href="#" onClick="window.open('../dalie/zj.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="官府放獵物了！">
<font color="#90CFEF">獵物</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fyuanbao.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="官府放元寶了！">
<font color="#90CFEF">元寶</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fjiangshi.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="官府放妖怪了！">
<font color="#90CFEF">妖怪</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fdog.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="官府放狗狗了！">
<font color="#90CFEF">小狗</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/dabianfu.ASP','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="官府放怪鳥了！">
<font color="#90CFEF">怪鳥</font></a>
<a class=blue href="#" onClick="window.open('../jiaofei/zj.asp','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="土匪出現了！">
<font color="#90CFEF">土匪</font></a>
<a class=blue href="#" onClick="window.open('guanadmin/fq.asp','fyuanbao','scrollbars=no,resizable=no,width=444,height=278')" title="發東西！">
<font color="#90CFEF">發東西</font></a>
<%else%>
<a href="../dt/reply.asp" target="_blank"><font color="#90CFEF">答題</font></a>
<a class=blue href="#" onClick="window.open('../dalie/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要官府放獵物了，你可以打獵物得錢！">
<font color="#90CFEF">打獵</font></a>
<a class=blue href="#" onClick="window.open('../jiaofei/fl.asp','daliwu','scrollbars=no,resizable=no,width=444,height=278')" title="只要有土匪了，你可以打土匪得東西了！">
<font color="#90CFEF">打土匪</font></a>
<%end if%>
<%if info(2)>=8 then%>
<a href="guanadmin/fine.asp" target="_blank"><font color="#90CFEF">抓人</font></a>
<%end if%>
<input type="hidden" name='towhoid' readonly value='0' size=2 maxlength=10>
<%if info(2)>=10 then%>
<a href="../linjianww/login.asp" target="_blank"><font color="#FF0000">高管</font>
<%end if%>
</a>
</div>
</td></tr>
</table><script language="JavaScript">function startnosay(){var nosay=99999;document.af.clock.value=nosay;}</script> 
<script src="readonly/f2.js"></script> <%if id=1 then%> <script>parent.rn();</script> 
<%else%> 
<script>parent.fc();</script> 
<%end if%>
<div id="dh" style="position:absolute; left:-80px; top:-80px; width:0px; height:0px; z-index:1"> 
<input type="button" value="<" name='gopp' onClick="Javascript:gop();" accesskey=","> 
<input type="button" value=">" name='gonn' onClick="Javascript:gon();" accesskey="."> 
</div></div></td></tr> </table></form></div>
</body></html>
