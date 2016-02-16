<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if Instr(LCase(useronlinename),LCase(" "&info(0)&" "))=0 then 
Session("ljjh_inthechat")="0" 
Response.Redirect "chaterr.asp?id=001" 
end if 
grade=Int(info(2))
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
if DateDiff("s",Session("ljjh_lasttime"),sj2)<2 then
Response.Write "<script language=JavaScript>{alert('有話慢慢說，別噎著哦！');}</script>"
Response.End
end if
t="..."
'是否被點穴
Application.Lock
if Instr(LCase(application("ljjh_dianxuename")),LCase(info(0)))>0 then
dianxue=split(application("ljjh_dianxuename"),";")
dxdx=0
for x=0 to UBound(dianxue)
dxname=split(dianxue(x),"|")
if dxname(0)=info(0) then
if DateDiff("s",dxname(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('你已被高手點了啞穴或被狼牙棒所傷，不能做任何事啊！還剩" & 60-DateDiff("s",dxname(1),sj2) & "秒自動解開!');</script>"
response.end
else
dxcz=dianxue(x)&";"
application("ljjh_dianxuename")=replace(application("ljjh_dianxuename"),dxcz,"")
Response.Write "<script Language=Javascript>alert('時間到了，你的啞穴自動解開了!');</script>"
response.end
end if
end if
next
end if
Application.UnLock
towho=Trim(Request.Form("towho"))
says=Trim(Request.Form("sy"))
says=replace(says," ","")
if towho="" then towho="大家"
if len(towho)>10 then towho=Left(towho,10)
gw=left(towho,1)
if IsNumeric(gw)=true then
zz=split(gw,"級")
gw=1
else 
gw=0
end if  
if Not(gw=1 or towho="大家" or towho="超級出氣包" or towho="戰神" or towho="狂殺大盜" or Instr(useronlinename," "&towho&" ")<>0 or Instr(useronlinename," "&nickmane&" ")<>0) then
Response.Write "<script Language=Javascript>alert('“" & towho & "”不在聊天室中，不能對其發言。');parent.f2.document.af.towho.options[0].value='大家';parent.f2.document.af.towho.options[0].text='大家';parent.f3.location.reload();</script>"
Response.end
end if
act=0
towhoway=Request.Form("towhoway")
if towhoway<>1 then towhoway=0
addwordcolor=Request.Form("addwordcolor")
sayscolor=Request.Form("sayscolor")
addsays=Request.Form("addsays")
towhoid=Trim(Request.Form("towhoid"))
titleline=Request.Form("titleline")
if InStr(addsays,"for(")<>0 or InStr(addsays,"alert(")<>0  then 
Response.Write "<script language=JavaScript>{alert('輸入有誤！');}</script>"
Response.End 
end if

if info(3)<20 or titleline<>1 then titleline=0
if titleline<>1 then titleline=0
if Instr(says,".")<>0 or Instr(says,"/")<>0 or Instr(says,"／")<>0 or Instr(says,"﹒")<>0 then titleline=0
if towho="大家" or titleline=1 then towhoway=0
if towho="江湖總站" then towhoway=1
if len(says)>150 then says=Left(says,150)
if (says="" or says="//") then Response.End
says=Replace(says,"&amp;","&")
//says=lcase(says)
says=Replace(says,"&amp;","&")
//says=lcase(says)
if info(2)<9 then
says=replace(says,"&#","")
says=replace(says,"..","")
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=Replace(says,"<"," ")
says=Replace(says,">"," ")
says=Replace(says,"\x3c"," ")
says=Replace(says,"\x3e"," ")
says=Replace(says,"\074"," ")
says=Replace(says,"\74"," ")
says=Replace(says,"\75"," ")
says=Replace(says,"\76"," ")
says=Replace(says,"&lt"," ")
says=Replace(says,"&gt"," ")
says=Replace(says,"\076"," ")
says=Replace(LCase(says),LCase("con"),"f2uck")
says=Replace(LCase(says),LCase("or"),"ｏｒ")
says=Replace(LCase(says),LCase("java"),"f2uck")
says=Replace(LCase(says),LCase("www"),"f2uck")
says=Replace(LCase(says),LCase("com"),"f2uck")
says=Replace(LCase(says),LCase("net"),"f2uck")
says=Replace(says,"game","f2uck")
says=Replace(says,"Ｇ","f2uck")
says=Replace(says,"asp","f2uck")
says=Replace(says,"d27","f2uck")
says=Replace(says,"66c","f2uck")
says=Replace(says,"６６","f2uck")
says=Replace(says,"００９","f2uck")
says=Replace(says,"009","f2uck")
says=Replace(says,"６６c","f2uck")
says=Replace(says,"jh","f2uck")
says=Replace(says,".to","f2uck")
says=Replace(LCase(says),LCase("http"),"f2uck")
maren=instr(says,"f2uck")
if maren<>0 then
Response.Write "<script Language=Javascript>top.location.href='autokick.asp';</script>"
Response.end
end if
end if
if info(2)<10 then
if instr(says,"發牌")=0 then 
if InStr(says,"加我")<>0 then Response.End
end if
end if
'是否戰鬥等級大於5才可以貼圖
if info(3)>2 and Instr(says," ")=0 and Instr(says,"[img]")<>0 then
if Instr(says,"[img]")<>0 then
if instr(says,"width")<>0 or instr(says,"height")<>0 then Response.End	
gsz=0
Do While Instr(says,"[img]")<>0
says=Replace(says,"[img]","<img src=pic/",1,1)
gsz=gsz+1
Loop
gsy=0
Do While Instr(says,"[/img]")<>0
says=Replace(says,"[/img]",">",1,1)
gsy=gsy+1
Loop
if gsz>gsy then says=says&">"
end if
end if
'虛擬加速設置xuni在//及/時取消虛擬顯示
'xuni=0
zj="<a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('[" & info(9) & "]');  target=f2><font color="& addwordcolor & ">" & info(0) & "</font></a>"
br="<a href=javascript:parent.sw('[" & towho & "]'); parent.sws('[" & towhoid & "]'); target=f2><font color=FFD7F4>" & towho & "</font></a>"
if Left(says,2)="//" then
'xuni=1
says=Replace(says,"##",zj,1,3,1)
says=Replace(says,"%%",br,1,3,1)
says="<font color=" & sayscolor & ">" & zj & Trim(mid(says,3,len(says)-2)) & "</font>"
act=1
titleline=0
end if

if left(says,1)="/" then
titleline=0
act=1
towhoway=0
neili=0
i=instr(says,"$")
if i=0 then i=len(says)+1
if trim(replace(mid(says,i+1)," ",""))="" and instr(";使用;怒吼;心跳;罰款;獎勵;好友;雙人賭博;清除;啞穴;點穴;釋放;發招;傳授;送錢;冊封;經驗;警告;加入;主題;心動;放大;存錢;取錢;轉賬;踢人;贈送;",right(left(says,i-1),len(left(says,i-1))-1))<>0 then
Response.Write "<script language=javascript>alert('【"&right(left(says,i-1),len(left(says,i-1))-1)&"】命令操作不能為空！');</script>"
Response.End
end if

fnn1=mid(says,i+1)
%><!--#include file="ljfunc/func.asp"--><%
select case left(says,i-1)
case "/主題","/title"
%><!--#include file="ljfunc/1.asp"--><%
says="<font color=" & sayscolor & ">"+titl(mid(says,i+1))+"</font>"
towho="大家"
towhoid=0
case "/領取會員費"
%><!--#include file="ljfunc/give1.asp"--><%
says="<font color=red>【領取會費】<font color=" & sayscolor & ">"+give1()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取金幣"
%><!--#include file="ljfunc/give2.asp"--><%
says="<font color=red>【領取金幣】<font color=" & sayscolor & ">"+give2()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取體力"
%><!--#include file="ljfunc/give3.asp"--><%
says="<font color=red>【領取體力】<font color=" & sayscolor & ">"+give3()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取內力"
%><!--#include file="ljfunc/give4.asp"--><%
says="<font color=red>【領取內力】<font color=" & sayscolor & ">"+give4()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取武功"
%><!--#include file="ljfunc/give5.asp"--><%
says="<font color=red>【領取武功】<font color=" & sayscolor & ">"+give5()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取攻擊"
%><!--#include file="ljfunc/give6.asp"--><%
says="<font color=red>【領取攻擊】<font color=" & sayscolor & ">"+give6()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取防御"
%><!--#include file="ljfunc/give7.asp"--><%
says="<font color=red>【領取防御】<font color=" & sayscolor & ">"+give7()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取魅力"
%><!--#include file="ljfunc/give8.asp"--><%
says="<font color=red>【領取魅力】<font color=" & sayscolor & ">"+give8()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取道德"
%><!--#include file="ljfunc/give9.asp"--><%
says="<font color=red>【領取道德】<font color=" & sayscolor & ">"+give9()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取法力"
%><!--#include file="ljfunc/give10.asp"--><%
says="<font color=red>【領取法力】<font color=" & sayscolor & ">"+give10()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/領取資質"
%><!--#include file="ljfunc/give01.asp"--><%
says="<font color=red>【領取資質】<font color=" & sayscolor & ">"+give01()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買會費"
%><!--#include file="ljfunc/kzd.asp"--><%
says="<font color=red>【購買會費】<font color=" & sayscolor & ">"+kzd()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買金幣"
%><!--#include file="ljfunc/kzda.asp"--><%
says="<font color=red>【購買金幣】<font color=" & sayscolor & ">"+kzda()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買子彈"
%><!--#include file="ljfunc/kzdb.asp"--><%
says="<font color=red>【購買子彈】<font color=" & sayscolor & ">"+kzdb()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買資質"
%><!--#include file="ljfunc/kzdc.asp"--><%
says="<font color=red>【購買資質】<font color=" & sayscolor & ">"+kzdc()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買法力"
%><!--#include file="ljfunc/kzdd.asp"--><%
says="<font color=red>【購買法力】<font color=" & sayscolor & ">"+kzdd()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買體力"
%><!--#include file="ljfunc/kzde.asp"--><%
says="<font color=red>【購買體力】<font color=" & sayscolor & ">"+kzde()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買武功"
%><!--#include file="ljfunc/kzdf.asp"--><%
says="<font color=red>【購買武功】<font color=" & sayscolor & ">"+kzdf()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買攻擊"
%><!--#include file="ljfunc/kzdg.asp"--><%
says="<font color=red>【購買攻擊】<font color=" & sayscolor & ">"+kzdg()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買防御"
%><!--#include file="ljfunc/kzdh.asp"--><%
says="<font color=red>【購買防御】<font color=" & sayscolor & ">"+kzdh()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買魅力"
%><!--#include file="ljfunc/kzdi.asp"--><%
says="<font color=red>【購買魅力】<font color=" & sayscolor & ">"+kzdi()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買道德"
%><!--#include file="ljfunc/kzdj.asp"--><%
says="<font color=red>【購買道德】<font color=" & sayscolor & ">"+kzdj()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買2級會員"
%><!--#include file="ljfunc/kzdk.asp"--><%
says="<font color=red>【購買2級會員】<font color=" & sayscolor & ">"+kzdk()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買3級會員"
%><!--#include file="ljfunc/kzdl.asp"--><%
says="<font color=red>【購買3級會員】<font color=" & sayscolor & ">"+kzdl()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買4級會員"
%><!--#include file="ljfunc/kzdm.asp"--><%
says="<font color=red>【購買4級會員】<font color=" & sayscolor & ">"+kzdm()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買6級官府"
%><!--#include file="ljfunc/kzdo.asp"--><%
says="<font color=red>【購買6級官府】<font color=" & sayscolor & ">"+kzdo()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買7級官府"
%><!--#include file="ljfunc/kzdp.asp"--><%
says="<font color=red>【購買7級官府】<font color=" & sayscolor & ">"+kzdp()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買8級官府"
%><!--#include file="ljfunc/kzdq.asp"--><%
says="<font color=red>【購買8級官府】<font color=" & sayscolor & ">"+kzdq()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買9級官府"
%><!--#include file="ljfunc/kzdr.asp"--><%
says="<font color=red>【購買9級官府】<font color=" & sayscolor & ">"+kzdr()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買10級官府"
%><!--#include file="ljfunc/kzds.asp"--><%
says="<font color=red>【購買10級官府】<font color=" & sayscolor & ">"+kzds()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買神龍戰士"
%><!--#include file="ljfunc/kzdt.asp"--><%
says="<font color=red>【購買神龍戰士】<font color=" & sayscolor & ">"+kzdt()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買金剛勇士"
%><!--#include file="ljfunc/kzdu.asp"--><%
says="<font color=red>【購買金剛勇士】<font color=" & sayscolor & ">"+kzdu()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/購買水天師"
%><!--#include file="ljfunc/kzdv.asp"--><%
says="<font color=red>【購買水天師】<font color=" & sayscolor & ">"+kzdv()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/點穴"
%><!--#include file="ljfunc/2.asp"--><%
says="<font color=00AFEF>【點穴】</font><font color=" & sayscolor & ">"+dian(towhoid,mid(says,i+1),towho)+"</font>"
case "/解穴"
%><!--#include file="ljfunc/3.asp"--><%
says="<font color=00AFEF>【解穴】</font><font color=" & sayscolor & ">"+jie(mid(says,i+1))+"</font>"
case "/逮捕"
%><!--#include file="ljfunc/4.asp"--><%
says="<font color=00AFEF>【逮捕】</font><font color=" & sayscolor & ">"+daipu(mid(says,i+1),towhoid,towho)+"</font>"
case "/坐牢"
%><!--#include file="ljfunc/5.asp"--><%
says="<font color=00AFEF>【坐牢】</font><font color=" & sayscolor & ">"+zuolao(mid(says,i+1),towhoid,towho)+"</font>"
case "/監禁"
%><!--#include file="ljfunc/6.asp"--><%
says="<font color=00AFEF>【監禁】</font><font color=" & sayscolor & ">"+jianjin(mid(says,i+1),towhoid,towho)+"</font>"
case "/釋放"
%><!--#include file="ljfunc/7.asp"--><%
says="<font color=00AFEF>【解除監禁】</font><font color=" & sayscolor & ">"+shifang(mid(says,i+1))+"</font>"
case "/警告","/掌門令"
%><!--#include file="ljfunc/9.asp"--><%
if info(2)<7 and info(6)="掌門" then
	says="<font color=00AFEF>【掌門令】<font color=" & sayscolor & ">"+jing(mid(says,i+1),towho)+"</font>"
else
	says="<font color=00AFEF>【警告】<font color=" & sayscolor & ">"+jing(mid(says,i+1),towho)+"</font>"
end if
case "/打架"
%><!--#include file="ljfunc/14c.asp"--><%
says="<font color=00AFEF>【打架】<font color=" & sayscolor & ">"+attackc(mid(says,i+1),towhoid)+"</font>"
case "/射擊"
%><!--#include file="ljfunc/14d.asp"--><%
says="<font color=00AFEF>【射擊】<font color=" & sayscolor & ">"+attackd(mid(says,i+1),towhoid,towho)+"</font>"
case "/下毒"
%><!--#include file="ljfunc/10.asp"--><%
says="<font color=00AFEF>【下毒】<font color=" & sayscolor & ">"+xiadu(mid(says,i+1),towhoid,towho)+"</font>"
case "/偷錢"
%><!--#include file="ljfunc/11.asp"--><%
says="<font color=00AFEF>【盜錢】<font color=" & sayscolor & ">"+touqian(towhoid,towho)+"</font>"
case "/賣了她吧"
%><!--#include file="ljfunc/53.asp"--><%
says="<font color=00AFEF>【拐賣行動】<font color=" & sayscolor & ">"+laobao(towhoid,towho)+"</font>"
case "/掃黃"
%><!--#include file="ljfunc/533.asp"--><%
says="<font color=00AFEF>【掃黃行動】<font color=" & sayscolor & ">"+saohuang(towhoid,towho)+"</font>"
case "/吸星大法"
%><!--#include file="ljfunc/12.asp"--><%
says="<font color=00AFEF>【吸星大法】<font color=" & sayscolor & ">"+xxdf(towhoid,towho)+"</font>"
case "/投擲"
%><!--#include file="ljfunc/13.asp"--><%
says="<font color=00AFEF>【發射暗器】<font color=" & sayscolor & ">"+touzi(mid(says,i+1),towhoid,towho)+"</font>"
case "/發招"
%><!--#include file="ljfunc/14.asp"--><%
says="<font color=00AFEF>【"&info(5)&"武學】<font color=" & sayscolor & ">"+attack(mid(says,i+1),towhoid,towho)+"</font>"
case "/搶奪"
%><!--#include file="ljfunc/14a.asp"--><%
says="<font color=00AFEF>【搶奪寶物】<font color=" & sayscolor & ">"+qiang(mid(says,i+1),towhoid,towho)+"</font>"
case "/傳送法力"
%><!--#include file="ljfunc/45002.asp"--><%
says="<font color=00AFEF>【傳送法力】<font color=" & sayscolor & ">"+songfali(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/尋找水晶"
%><!--#include file="ljfunc/45003.asp"--><%
says="<font color=green>[尋找水晶]<font color=" & sayscolor & ">"+xunshuijing(mid(says,i+1),towhoid,towho)+"</font>"
case "/尋找法器"
%><!--#include file="ljfunc/45005.asp"--><%
says="<font color=green>[尋找法器]<font color=" & sayscolor & ">"+xunfaqi(mid(says,i+1),towhoid,towho)+"</font>"
case "/尋找流星"
%><!--#include file="ljfunc/4000.asp"--><%
says="<font color=green>[尋找流星]<font color=" & sayscolor & ">"+bbbbbb(mid(says,i+1),towhoid,towho)+"</font>"
case "/尋找流星淚"
%><!--#include file="ljfunc/4001.asp"--><%
says="<font color=green>[尋找流星淚]<font color=" & sayscolor & ">"+cccccc(mid(says,i+1),towhoid,towho)+"</font>"
case "/尋找龍珠"
%><!--#include file="ljfunc/4002.asp"--><%
says="<font color=green>[尋找龍珠]<font color=" & sayscolor & ">"+dddddd(mid(says,i+1),towhoid,towho)+"</font>"
case "/尋找會費"
%><!--#include file="ljfunc/9000.asp"--><%
says="<font color=green>[尋找會費]<font color=" & sayscolor & ">"+dog(mid(says,i+1),towhoid,towho)+"</font>"
case "/尋找金幣"
%><!--#include file="ljfunc/9001.asp"--><%
says="<font color=green>[尋找金幣]<font color=" & sayscolor & ">"+dogdog(mid(says,i+1),towhoid,towho)+"</font>"
case "/魔力鑽石"
%><!--#include file="ljfunc/45006.asp"--><%
says="<font color=00AFEF>【魔力鑽石】<font color=" & sayscolor & ">"+molishi(mid(says,i+1),towhoid)+"</font>"
case "/生日蛋糕"
%><!--#include file="ljfunc/450018.asp"--><%
says="<font color=00AFEF>【生日蛋糕】<font color=" & sayscolor & ">"+shendangao(mid(says,i+1),towhoid)+"</font>"
case "/配制寶石"
%><!--#include file="ljfunc/45015.asp"--><%
says="<font color=00AFEF>【配制寶石】<font color=" & sayscolor & ">"+peibashi(mid(says,i+1),towhoid)+"</font>"
case "/發射子彈"
%><!--#include file="ljfunc/450017.asp"--><%
says="<font color=00AFEF>【發射子彈】<font color=" & sayscolor & ">"+fashezi(mid(says,i+1),towhoid,towho)+"</font>"
case "/傳授"
%><!--#include file="ljfunc/15.asp"--><%
says="<font color=00AFEF>【傳授】<font color=" & sayscolor & ">"+cuan(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/跟蹤私聊"
%><!--#include file="ljfunc/888.asp"--><%
says="<font color=00AFEF>【跟蹤私聊】<font color=" & sayscolor & ">"+tracksl()+"</font>" 
towhoway=1
towho=info(0)
towhoid=info(9)
case "/取消跟蹤" 
%><!--#include file="ljfunc/999.asp"--><%
says="<font color=00AFEF>【取消跟蹤】<font color=" & sayscolor & ">"+clearsl()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/禁打開關"
%><!--#include file="ljfunc/666.asp"--><%
says="<font color=00AFEF>【禁打開關】<font color=" & sayscolor & ">"+jingda()+"</font>" 
case "/隱身"
%><!--#include file="ljfunc/ying.asp"--><%
says="<font color=00AFEF>【隱身開關】<font color=" & sayscolor & ">"+yingsheng()+"</font>" 
towhoway=1
towho=info(0)
towhoid=info(9)
case "/送錢"
%><!--#include file="ljfunc/16.asp"--><%
says="<font color=00AFEF>【送錢】<font color=" & sayscolor & ">"+give(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/送會費"
%><!--#include file="ljfunc/2016.asp"--><%
says="<font color=00AFEF>【贈送會費】<font color=" & sayscolor & ">"+eeee(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/送金幣"
%><!--#include file="ljfunc/2017.asp"--><%
says="<font color=00AFEF>【贈送金幣】<font color=" & sayscolor & ">"+ffff(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/門規處置"
%><!--#include file="ljfunc/161.asp"--><%
says="<font color=00AFEF>【門規處置】<font color=" & sayscolor & ">"+mengui(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/罰款"
%><!--#include file="ljfunc/611.asp"--><%
says="<font color=00AFEF>【江湖罰款】</font><font color=" & sayscolor & ">"+fakuan(mid(says,i+1),towhoid,towho)+"</font>"
case "/強制出關術"
%><!--#include file="ljfunc/2000.asp"--><%
says="<font color=00AFEF>【強制出關術】</font><font color=" & sayscolor & ">"+aa(mid(says,i+1),towhoid,towho)+"</font>"
case "/地獄減管術"
%><!--#include file="ljfunc/2001.asp"--><%
says="<font color=00AFEF>【地獄減管術】</font><font color=" & sayscolor & ">"+bb(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加管術"
%><!--#include file="ljfunc/2002.asp"--><%
says="<font color=00AFEF>【光明加管術】</font><font color=" & sayscolor & ">"+cc(mid(says,i+1),towhoid,towho)+"</font>"
case "/黑暗減攻術"
%><!--#include file="ljfunc/2003.asp"--><%
says="<font color=00AFEF>【黑暗減攻術】</font><font color=" & sayscolor & ">"+dd(mid(says,i+1),towhoid,towho)+"</font>"
case "/黑暗減防術"
%><!--#include file="ljfunc/2004.asp"--><%
says="<font color=00AFEF>【黑暗減防術】</font><font color=" & sayscolor & ">"+ee(mid(says,i+1),towhoid,towho)+"</font>"
case "/黑暗減員術"
%><!--#include file="ljfunc/2005.asp"--><%
says="<font color=00AFEF>【黑暗減員術】</font><font color=" & sayscolor & ">"+ff(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加員術"
%><!--#include file="ljfunc/2006.asp"--><%
says="<font color=00AFEF>【光明加員術】</font><font color=" & sayscolor & ">"+aaa(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加防術"
%><!--#include file="ljfunc/2007.asp"--><%
says="<font color=00AFEF>【光明加防術】</font><font color=" & sayscolor & ">"+bbb(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加攻術"
%><!--#include file="ljfunc/2008.asp"--><%
says="<font color=00AFEF>【光明加攻術】</font><font color=" & sayscolor & ">"+ccc(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加錢術"
%><!--#include file="ljfunc/5001.asp"--><%
says="<font color=00AFEF>【光明加錢術】</font><font color=" & sayscolor & ">"+abcabc(mid(says,i+1),towhoid,towho)+"</font>"
case "/黑暗減金術"
%><!--#include file="ljfunc/2009.asp"--><%
says="<font color=00AFEF>【黑暗減金術】</font><font color=" & sayscolor & ">"+ddd(mid(says,i+1),towhoid,towho)+"</font>"
case "/黑暗減會術"
%><!--#include file="ljfunc/2010.asp"--><%
says="<font color=00AFEF>【黑暗減會術】</font><font color=" & sayscolor & ">"+eee(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加金術"
%><!--#include file="ljfunc/2011.asp"--><%
says="<font color=00AFEF>【光明加金術】</font><font color=" & sayscolor & ">"+fff(mid(says,i+1),towhoid,towho)+"</font>"
case "/光明加會術"
%><!--#include file="ljfunc/2012.asp"--><%
says="<font color=00AFEF>【光明加會術】</font><font color=" & sayscolor & ">"+aaaa(mid(says,i+1),towhoid,towho)+"</font>"
case "/耿鬼催眠術"
%><!--#include file="ljfunc/2013.asp"--><%
says="<font color=00AFEF>【耿鬼催眠術】</font><font color=" & sayscolor & ">"+bbbb(mid(says,i+1),towhoid,towho)+"</font>"
case "/富迪解眠術"
%><!--#include file="ljfunc/2014.asp"--><%
says="<font color=00AFEF>【富迪解眠術】</font><font color=" & sayscolor & ">"+cccc(mid(says,i+1),towhoid,towho)+"</font>"
case "/地獄減錢術"
%><!--#include file="ljfunc/5002.asp"--><%
says="<font color=00AFEF>【地獄減錢術】</font><font color=" & sayscolor & ">"+gogo(mid(says,i+1),towhoid,towho)+"</font>"
case "/斬首術"
%><!--#include file="ljfunc/2015.asp"--><%
says="<font color=00AFEF>【斬首術】</font><font color=" & sayscolor & ">"+dddd(mid(says,i+1),towhoid,towho)+"</font>"
case "/門規處置"
%><!--#include file="ljfunc/161.asp"--><%
says="<font color=00AFEF>【門規處置】<font color=" & sayscolor & ">"+mengui(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/罰款"
%><!--#include file="ljfunc/611.asp"--><%
says="<font color=00AFEF>【江湖罰款】</font><font color=" & sayscolor & ">"+fakuan(mid(says,i+1),towhoid,towho)+"</font>"
case "/動刑"
%><!--#include file="ljfunc/612.asp"--><%
says="<font color=00AFEF>【江湖動刑】</font><font color=" & sayscolor & ">"+dongxin(mid(says,i+1),towhoid,towho)+"</font>"
case "/贈送"
%><!--#include file="ljfunc/17.asp"--><%
says="<font color=00AFEF>【贈送】<font color=" & sayscolor & ">"+zen(mid(says,i+1),towhoid,towho)+"</font>"
case "/丟棄"
%><!--#include file="ljfunc/20031.asp"--><%
says="<font color=00AFEF>【丟棄物品】<font color=" & sayscolor & ">"+diu(mid(says,i+1),towhoid,towho)+"</font>"
case "/加入"
%><!--#include file="ljfunc/19.asp"--><%
says="<font color=00AFEF>【加入】<font color=" & sayscolor & ">"+join(mid(says,i+1))+"</font>"
case "/冊封"
%><!--#include file="ljfunc/20.asp"--><%
says="<font color=00AFEF>【冊封】<font color=" & sayscolor & ">"+cefen(mid(says,i+1),towhoid,towho)+"</font>"
case "/公告"
%><!--#include file="ljfunc/21.asp"--><%
says=gong(mid(says,i+1))
case "/踢人"
%><!--#include file="ljfunc/22.asp"--><%
says="<font color=00AFEF>【踢人】<font color=" & sayscolor & ">"+tiren(towhoid,towho,mid(says,i+1))+"</font>"
case "/封鎖ip"
%><!--#include file="ljfunc/22001.asp"--><%
says="<font color=00AFEF>【封鎖ip】<font color=" & sayscolor & ">"+jiafeng(towhoid,towho,mid(says,i+1))+"</font>"
case "/解鎖ip"
%><!--#include file="ljfunc/22002.asp"--><%
says="<font color=00AFEF>【解鎖ip】</font><font color=" & sayscolor & ">"+jiansu(mid(says,i+1))+"</font>"
case "/心動"
%><!--#include file="ljfunc/23.asp"--><%
says=xindong(mid(says,i+1))
towho="大家"
towhoid=0
case "/放大"
%><!--#include file="ljfunc/7010.asp"--><%
says=rock(mid(says,i+1))
towho="大家"
towhoid=0
case "/藍光字"
%><!--#include file="ljfunc/24.asp"--><%
says=fangda(mid(says,i+1))
towho="大家"
towhoid=0
case "/紅光字"
%><!--#include file="ljfunc/1000.asp"--><%
says=a(mid(says,i+1))
towho="大家"
towhoid=0
case "/綠光字"
%><!--#include file="ljfunc/1001.asp"--><%
says=b(mid(says,i+1))
towho="大家"
towhoid=0
case "/黃光字"
%><!--#include file="ljfunc/1002.asp"--><%
says=c(mid(says,i+1))
towho="大家"
towhoid=0
case "/紫光字"
%><!--#include file="ljfunc/1003.asp"--><%
says=d(mid(says,i+1))
towho="大家"
towhoid=0
case "/底色文字"
%><!--#include file="ljfunc/1004.asp"--><%
says=e(mid(says,i+1))
towho="大家"
towhoid=0
case "/拜師"
%><!--#include file="ljfunc/25.asp"--><%
says="<font color=00AFEF>【拜師】<font color=" & sayscolor & ">"+bais(towhoid,towho)+"</font>"
case "/下山"
%><!--#include file="ljfunc/251.asp"--><%
says="<font color=00AFEF>【下山】<font color=" & sayscolor & ">"+xiashan(towhoid,towho)+"</font>"
case "/收徒"
%><!--#include file="ljfunc/26.asp"--><%
says="<font color=00AFEF>【收徒】<font color=" & sayscolor & ">"+stu(towho)+"</font>"
case "/求妾","/求小老公"
%><!--#include file="ljfunc/00.asp"--><%
says="<font color=00AFEF>【求妾】<font color=" & sayscolor & ">"+qqq(towhoid,towho)+"</font>"
case "/同意"
%><!--#include file="ljfunc/000.asp"--><%
says="<font color=00AFEF>【同意】<font color=" & sayscolor & ">"+dy(towho)+"</font>"
case "/單挑取消"
%><!--#include file="ljfunc/311.asp"--><%
says="<font color=00AFEF>【單挑取消】</font><font color=" & sayscolor & ">"+quxiaoa()+"</font>"
case "/打坐"
%><!--#include file="ljfunc/27.asp"--><%
says="<font color=red>【打坐】<font color=" & sayscolor & ">"+dazhuo()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/閉目"
%><!--#include file="ljfunc/28.asp"--><%
says="<font color=red>【閉目】<font color=" & sayscolor & ">"+bimu()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/修習武功"
%><!--#include file="ljfunc/71.asp"--><%
says="<font color=00AFEF>【修習武功】</font><font color=" & sayscolor & ">"+wuxiu(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/參禪"
%><!--#include file="ljfunc/canchan.asp"--><%
says="<font color=red>【山洞參禪】<font color=" & sayscolor & ">"+canchan()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/神思"
%><!--#include file="ljfunc/shenci.asp"--><%
says="<font color=red>【靈劍神思】<font color=" & sayscolor & ">"+shenci()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/修養"
%><!--#include file="ljfunc/xiuyang.asp"--><%
says="<font color=red>【修心養性】<font color=" & sayscolor & ">"+xiuyang()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/經驗"
%><!--#include file="ljfunc/29.asp"--><%
says="<font color=red>【傳送經驗】<font color=" & sayscolor & ">"+jingyan(mid(says,i+1),towhoid,towho)+"</font>"
case "/練武"
%><!--#include file="ljfunc/30.asp"--><%
says="<font color=red>【練武】<font color=" & sayscolor & ">"+lianwu()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/拼命"
%><!--#include file="ljfunc/59.asp"--><%
says="<font color=00AFEF>【同歸於盡】<font color=" & sayscolor & ">"+pingmin(mid(says,i+1),towhoid,towho)+"</font>"
case "/醫療"
%><!--#include file="ljfunc/61.asp"--><%
says="<font color=red>【妙手回春】<font color=" & sayscolor & ">"+yiliao()+"</font>"
case "/美容"
%><!--#include file="ljfunc/62.asp"--><%
says="<font color=red>【美容大法】<font color=" & sayscolor & ">"+meirong()+"</font>"
case "/教武"
%><!--#include file="ljfunc/63.asp"--><%
says="<font color=red>【教導武功】<font color=" & sayscolor & ">"+junguan()+"</font>"
case "/補氣"
%><!--#include file="ljfunc/64.asp"--><%
says="<font color=red>【遊雲靈氣】<font color=" & sayscolor & ">"+qigongshi()+"</font>"

case "/存錢"
%><!--#include file="ljfunc/31.asp"--><%
says="<font color=00AFEF>【存錢】<font color=" & sayscolor & ">"+putyin(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/取錢"
%><!--#include file="ljfunc/32.asp"--><%
says="<font color=00AFEF>【取錢】<font color=" & sayscolor & ">"+getyin(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/存法力"
%><!--#include file="ljfunc/450020.asp"--><%
says="<font color=00AFEF>【存法力】<font color=" & sayscolor & ">"+putfali(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/取法力"
%><!--#include file="ljfunc/450021.asp"--><%
says="<font color=00AFEF>【取法力】<font color=" & sayscolor & ">"+getfali(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/轉賬"
%><!--#include file="ljfunc/33.asp"--><%
says="<font color=00AFEF>【轉賬】<font color=" & sayscolor & ">"+zzyin(mid(says,i+1),towhoid,towho)+"</font>"
case "/修練"
%><!--#include file="ljfunc/35.asp"--><%
says="<font color=red>【寶物修練】<font color=" & sayscolor & ">"+xiulian()+"</font>"
case "/暴豆"
%><!--#include file="ljfunc/36.asp"--><%
says="<font color=red>【暴威力豆】<font color=" & sayscolor & ">"+baodou()+"</font>"
case "/啞穴"
%><!--#include file="ljfunc/37.asp"--><%
says="<font color=00AFEF>【啞穴】<font color=" & sayscolor & ">"+ya(towhoid,towho,mid(says,i+1))+"</font>"
case "/斬首"
%><!--#include file="ljfunc/8.asp"--><%
says="<font color=00AFEF>【斬首】<font color=" & sayscolor & ">"+zhanshou(mid(says,i+1),towhoid)+"</font>"
case "/怒吼"
%><!--#include file="ljfunc/39.asp"--><%
says=nuhou(mid(says,i+1))
towho="大家"
case "/心跳"
%><!--#include file="ljfunc/40.asp"--><%
says=xintiao(mid(says,i+1))
towho="大家"
case "/查ip"
%><!--#include file="ljfunc/41.asp"--><%
says=seeip(towhoid,towho)
towhoway=1
towho=info(0)
case "/獎勵"
%><!--#include file="ljfunc/42.asp"--><%
says="<font color=00AFEF>【獎勵】<font color=" & sayscolor & ">"+jiangli(mid(says,i+1),towhoid,towho)+"</font>"
case "/官府獎勵"
%><!--#include file="ljfunc/jianglila.asp"--><%
says="<font color=00AFEF>【官府獎勵】<font color=" & sayscolor & ">"+jianglila(mid(says,i+1),towhoid,towho)+"</font>"
case "/基金"
%><!--#include file="ljfunc/43.asp"--><%
says="<font color=red>【基金查看】<font color=" & sayscolor & ">"+seejj()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/使用"
%><!--#include file="ljfunc/44.asp"--><%
says="<font color=00AFEF>【卡片】<font color=" & sayscolor & ">"+kapian(mid(says,i+1),towhoid,towho)+"</font>"
case "/小偷"
%><!--#include file="ljfunc/45.asp"--><%
says="<font color=00AFEF>【小偷】<font color=" & sayscolor & ">"+tdx(towhoid,towho)+"</font>"
case "/抓小偷"
%><!--#include file="ljfunc/46.asp"--><%
says="<font color=00AFEF>【抓小偷】<font color=" & sayscolor & ">"+pk(towhoid,towho)+"</font>"
case "/孩子偷竊"
%><!--#include file="ljfunc/51.asp"--><%
says="<font color=00AFEF>【孩子偷竊】<font color=" & sayscolor & ">"+hza(towhoid,towho)+"</font>"
case "/教訓孩子"
%><!--#include file="ljfunc/52.asp"--><%
says="<font color=00AFEF>【教訓孩子】<font color=" & sayscolor & ">"+hzb(towhoid,towho)+"</font>"
case "/雙人賭博"
%><!--#include file="ljfunc/54.asp"--><%
says="<font color=00AFEF>【雙人賭博】<font color=" & sayscolor & ">"+grdb(mid(says,i+1),towhoid,towho)+"</font>"
case "/求婚"
%><!--#include file="ljfunc/47.asp"--><%
says="<font color=00AFEF>【求婚】<font color=" & sayscolor & ">"+qiuhun(mid(says,i+1),towhoid,towho)+"</font>"
case "/同盟"
%><!--#include file="ljfunc/jiemen.asp"--><%
says="<font color=00AFEF>【幫派結盟】<font color=" & sayscolor & ">"+jiemen(towhoid,towho)+"</font>"
case "/解除同盟"
%><!--#include file="ljfunc/jiecmen.asp"--><%
says="<font color=00AFEF>【解除同盟】<font color=" & sayscolor & ">"+jiecmen(mid(says,i+1),towhoid,towho)+"</font>"
case "/幫派大戰"
%><!--#include file="ljfunc/bangpai.asp"--><%
says="<font color=00AFEF>【幫派大戰】<font color=" & sayscolor & ">"+bangpai(towhoid,towho)+"</font>"
case "/打撲克"
%><!--#include file="duju/dpk-ask.asp"--><%
says="<font color=00AFEF>【打撲克】<font color=" & sayscolor & ">"+dpkask(mid(says,i+1),towhoid,towho)+"</font>"
case "/發牌"
%><!--#include file="duju/dpkfp.asp"--><%
says="<font color=00AFEF>【發牌】<font color=" & sayscolor & ">"+dpkfp(mid(says,i+1),towhoid,towho)+"</font>"
case "/打麻將"
%><!--#include file="duju/dmj-ask.asp"--><%
says="<font color=00AFEF>【打麻將】<font color=" & sayscolor & ">"+dmjask(mid(says,i+1),towhoid,towho)+"</font>"
case "/出牌"
%><!--#include file="duju/dmjfp.asp"--><%
says="<font color=00AFEF>【出牌】<font color=" & sayscolor & ">"+dmjfp(mid(says,i+1),towhoid,towho)+"</font>"
case "/摸牌"
%><!--#include file="duju/dmjmp.asp"--><%
says="<font color=00AFEF>【摸牌】<font color=" & sayscolor & ">"+dmjGetmj(towhoid,towho)+"</font>"
case "/喫牌"
%><!--#include file="duju/dmjcp.asp"--><%
says="<font color=00AFEF>【喫牌】<font color=" & sayscolor & ">"+dmjPeng(mid(says,i+1),towhoid,towho)+"</font>"
case "/踫牌"
%><!--#include file="duju/dmjcp.asp"--><%
says="<font color=00AFEF>【踫牌】<font color=" & sayscolor & ">"+dmjPeng(mid(says,i+1),towhoid,towho)+"</font>"
case "/問牌"
%><!--#include file="duju/dmjwp.asp"--><%
says="<font color=00AFEF>【問牌】<font color=" & sayscolor & ">"+dmjwp(towho)+"</font>"
case "/單挑"
%><!--#include file="ljfunc/471.asp"--><%
says="<font color=red>【生死決鬥】<font color=" & sayscolor & ">"+diantiao(mid(says,i+1),towhoid,towho)+"</font>"
case "/招收"
%><!--#include file="ljfunc/4711.asp"--><%
says="<font color=red>【廣招弟子】<font color=" & sayscolor & ">"+jiamen(mid(says,i+1),towhoid,towho)+"</font>"
case "/幸運"
%><!--#include file="ljfunc/48.asp"--><%
says="<font color=00AFEF>【幸運風采】<font color=" & sayscolor & ">"+xingyu(mid(says,i+1))+"</font>"
if instr(says,"恭喜")=0 then
towhoway=1
towho=info(0)
towhoid=info(9)
end if
case "/狀態"
%><!--#include file="ljfunc/18.asp"--><%
says="<font color=00AFEF>【狀態】<font color=" & sayscolor & ">"&info(0)&"得到江湖情報:"+getstat(towhoid,towho)+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/心情"
%><!--#include file="ljfunc/55.asp"--><%
says="<font color=00AFEF>【心情】<font color=" & sayscolor & ">"+xinqing(mid(says,i+1))+"</font>"
towho="mingdan"
case "/分類物品"
Response.Write "<script language=javascript>window.open('flwp.asp?fl="& mid(says,i+1) &"');</script>"
Response.End
case else
Response.Write "<script Language=Javascript>alert('您輸入的命令錯誤！');</script>"
response.end
end select
nowcz=0
if len(info(0))<=len(towho)  then
if instr(towho,info(0))>0 then nowcz=1
end if
if nowcz=0 then
says=Replace(says,info(0),"<a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('[" & info(9) & "]'); target=f2  ><font  color="&addwordcolor&">"& info(0) &"</font></a>")
end if
end if
addsays=addsays&"對"
Session("ljjh_lasttime")=sj2
says=Replace(says,"\","\\")
says=Replace(says,"/","\/")
says=Replace(says,chr(34),"\"&chr(34))
if titleline=1 then
Application.Lock
Application("ljjh_title"&session("nowinroom"))=Left(says,32) & "<font color=00FFFF style=font-size:11pt>（" & info(0) & "，" & sj & "）</font>"
Application.UnLock
call showsay()
end if
Application.Lock
	sd=Application("ljjh_sd")
	line=int(Application("ljjh_line"))
	Application("ljjh_line")=line+1
Application.UnLock
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
if act=1 then
sd(192)=-1
else
sd(192)=info(9)
end if
sd(193)=towhoway
sd(194)=info(0)
sd(195)=towho
sd(196)=addwordcolor
sd(197)=sayscolor
sd(198)=addsays
sd(199)=says&"..."
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
if chatinfo(6)<>0 then
randomize()
rnd1=int(rnd*750)
sayyq=""
if rnd1<=2 then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
conn.execute "Update 用戶 set 保護=false where 會員等級=0"
conn.close
set conn=nothing
sayyq="<center><font size=4 color=red>【一陣大風刮來,所有的非會員都被解除保護了..怕死的快開保護阿!】</center><bgsound src=wav/003.MID loop=1></font>"
end if
if rnd1>=21 and rnd1<=25 then
abc="<a href='qiang.asp' target='d'><img src='img/tank.gif' border='0'></a>"
sayyq="<font color=red>一把精制手槍出現在聊天室裡，大家快搶啊!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
end if
if rnd1>=26 and rnd1<=30 then
Application.Lock
Application("ljjh_jiaofei")="土匪"
Application.UnLock
sayyq="<font color=red>【消息】</font><font color=red>突然一群土匪<img src=img/jiao.gif>闖入江湖搶劫，請高手們快去剿匪啊。</font>"
end if
if rnd1>=31 and rnd1<=35 then
jstl=int(rnd*5000)+100
Application.Lock
Application("ljjh_kl")=jstl
Application.UnLock
abc="<a href='kl.asp' target='d'><img src='img/shenmo.GIF' border='0' width='54' height='54'></a>"
sayyq="<font color=red>突然間妖風陣陣  “妖怪呀！”一女子尖叫，一頭妖怪闖進聊天室，妖怪體力：+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
end if
	if sayyq<>"" then
		Application.Lock
		sd=Application("ljjh_sd")
		line=int(Application("ljjh_line"))
		Application("ljjh_line")=line+1
		Application.UnLock
		for i=1 to 190
		sd(i)=sd(i+10)
		next
		sd(191)=line+1
		sd(192)=-1
		sd(193)=0
		sd(194)=info(0)
		sd(195)="大家"
		sd(196)=addwordcolor
		sd(197)=sayscolor
		sd(198)=addsays
		sd(199)=sayyq&"..."
		sd(200)=session("nowinroom")
		Application.Lock
		Application("ljjh_sd")=sd
		Application.UnLock
	end if
end if
call showsay()
Sub showsay()
filname=Session("ljjh_filname")
Response.Write "<script Language=JavaScript>"
sd=Application("ljjh_sd")
userline=int(Session("ljjh_line"))
newuserline=0
Dim show()
j=1
for i=1 to 200 step 10
newuserline=sd(i)
if sd(i)>userline and cstr(sd(i+9))=cstr(session("nowinroom")) and (session("slshow")=1 or (sd(i+2)="0" or sd(i+4)="大家" or sd(i+2)="1" and (CStr(sd(i+3))=CStr(info(0)) or CStr(sd(i+4))=CStr(info(0)))) and Instr(filname," "&sd(i+3)&",")=0) then
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5),show(j+6),show(j+7),show(j+8),show(j+9)
show(j)=sd(i)
show(j+1)=sd(i+1)
show(j+2)=sd(i+2)
show(j+3)=sd(i+3)
show(j+4)=sd(i+4)
show(j+5)=sd(i+5)
show(j+6)=sd(i+6)
show(j+7)=sd(i+7)
show(j+8)=sd(i+8)
show(j+9)=sd(i+9)
j=j+10
end if
next
if j>1  then
		for i=1 to UBound(show) step 10
		Response.Write "parent.sh(" & show(i+1) & ","& towhoid & "," & show(i+2) & "," & chr(34) & show(i+3) & chr(34) & "," & chr(34) & show(i+4) & chr(34) & "," & chr(34) & show(i+5) & chr(34) & "," & chr(34) & show(i+6) & chr(34) & "," & chr(34) & show(i+7) & chr(34) & "," & chr(34) & show(i+8) & chr(34) & ");"
		next
end if
if titleline=1 then Response.Write "parent.t.location.reload();"
Response.Write "</script>"
if newuserline>userline then Session("ljjh_line")=newuserline
Response.End
End Sub
%>