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
says="<fo