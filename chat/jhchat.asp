<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_zm= info(4) & "|" & info(5) & "|"&info(0)&"|"&info(1)&"|"&info(12)&"|" & info(6) &"|"&info(3)&"|"&info(9)&"|"&info(10)&"|"&info(8)&"|"&info(14)
scrname=Request.ServerVariables("SCRIPT_NAME")
if Instr(LCase(scrname),"jhchat.asp")=0 then Response.Redirect "../error.asp?id=002"
Session("ljjh_filname")=" ,"
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
'這一行是檢測用戶是否為死或是基他的
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 操作時間,狀態,配偶 FROM 用戶 WHERE 狀態='正常' and id="&info(9)
rs.open sql,conn,1,3
if rs.Eof and rs.Bof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../error.asp?id=456"
end if
if Session("ljjh_inthechat")<>"1" then
conn.execute "update 用戶 set 操作時間=now() where id="&info(9)
end if
zt=rs("狀態")
pe=rs("配偶")
rs.close
set rs=nothing
conn.close
set conn=nothing

if zt="" or zt="死" or zt="眠" or zt="無" or zt="牢" or zt="獄" or zt="監禁" then  Response.Redirect "../error.asp?id=456"
do while DateDiff("s",Application("ljjh_czsj"),now())<3
loop
Application.Lock
Application("ljjh_czsj")=now()
Application.UnLock
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t=s & ":" & f & ":" & m
Session("ljjh_lastsaytime")=sj
if Session("ljjh_inthechat")<>"1" then
for i=0 to chatroomnum
if InStr(LCase(Application("ljjh_useronlinename"&i))," " & LCase(info(0)) & " ")<>0 then Response.Redirect "../error.asp?id=300"
next
Session("ljjh_inthechat")="1"
Session("ljjh_savetime")=now()
Session("ljjh_lasttime")=sj
Application.Lock
dim newonlinelist()
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
useronlinename=""
onliners=0
js=1
yjl=0
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if DateDiff("n",onlinelist(i+5),sj)<=60 then
if yjl=0 and len(onlinelist(i+1))>len(info(0)) then
yjl=1
onliners=onliners+2
useronlinename=useronlinename & " " & info(0) & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5),newonlinelist(js+6),newonlinelist(js+7),newonlinelist(js+8),newonlinelist(js+9),newonlinelist(js+10),newonlinelist(js+11)
newonlinelist(js)=info(9)
newonlinelist(js+1)=info(0)
newonlinelist(js+2)=info(1)&info(4)
newonlinelist(js+3)=ljjh_zm
newonlinelist(js+4)=sj
newonlinelist(js+5)=sj
newonlinelist(js+6)=onlinelist(i)
newonlinelist(js+7)=onlinelist(i+1)
newonlinelist(js+8)=onlinelist(i+2)
newonlinelist(js+9)=onlinelist(i+3)
newonlinelist(js+10)=onlinelist(i+4)
newonlinelist(js+11)=onlinelist(i+5)
js=js+12
else
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
end if
end if
next
if yjl=0 then
onliners=onliners+1
useronlinename=useronlinename & " " & info(0)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=info(9)
newonlinelist(js+1)=info(0)
newonlinelist(js+2)=info(1)&info(4)
newonlinelist(js+3)=ljjh_zm
newonlinelist(js+4)=sj
newonlinelist(js+5)=sj
end if
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
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Session("ljjh_line")=line

'站長進入是否顯示信息，如果不顯示改成'不讓執行
if Instr(1,Application("ljjh_ying"),"|"& info(0) & "|")=0  then
Application.Lock
Application("ljjh_line")=line+1
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)=info(0)
sd(195)="mingdan"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
randomize timer
chatdj=int(rnd*9)+1
aszcc=int(rnd*9)+1
Select Case chatdj
case 1
ljjh_userinto="在江湖的大道上,%%騎著一頭未滿周歲的豬娃，手拿一根柳樹條，來到了"&Application("ljjh_chatroomname")&",“衝呀、殺呀”的,急忙拱手曰:“各位兄長，小弟有禮!”"
case 2
ljjh_userinto="在江湖的大道上,%%騎著自己家的狗兒大黑子，手拿一把大菜刀來到了"&Application("ljjh_chatroomname")&"，“哼、我怕誰呀！”見到了江湖眾位兄弟,急忙拱手曰:“，各位兄長，小弟有禮!”"
case 3
ljjh_userinto="在江湖的大道上,%%騎著一頭黑毛驢哼著小曲，腰間一把生了繡的破鐵劍，來到了"&Application("ljjh_chatroomname")&",“江湖路難行呀！”，拱手曰:“各位兄弟，小生有禮!”"
case 4
ljjh_userinto="在江湖的大道上,%%騎著一匹瘸腳老馬，腰間一把大刀，來到了"&Application("ljjh_chatroomname")&",“江湖任我行呀！哈。。”，拱手曰:“各位兄弟，小生有禮!”"
case 5
ljjh_userinto="在江湖的大道上,%%騎著一匹高頭大馬，腰間一把王八盒子（老式手槍），來到了"&Application("ljjh_chatroomname")&",“呵。。總算現代化了！”，拱手曰:“各位兄弟，小生有禮!”"
case 6
ljjh_userinto="在江湖的大道上,%%騎著一輛國產破踏板（黑煙直冒），腰間一把五4，來到了"&Application("ljjh_chatroomname")&",“呵。。老子也有今天！”，拱手曰:“各位兄弟，小生有禮!”"
case 7
ljjh_userinto="在江湖的大道上,%%開著一輛夏麗（熱的呼呼呼直喘氣），腰間一把半自動，來到了"&Application("ljjh_chatroomname")&",“呵。。我有福了！”，拱手曰:“各位兄弟，小生有禮!”"
case 8
ljjh_userinto="在江湖的大道上,%%開著一輛夏麗（熱的呼呼呼直喘氣），腰間一把半自動，來到了"&Application("ljjh_chatroomname")&",“呵。。我有福了！”，拱手曰:“各位兄弟，小生有禮!”"
case 9
ljjh_userinto="在江湖的大道上,%%開著一輛奧迪（真是舒服），腰間一把全自動，來到了"&Application("ljjh_chatroomname")&",“呵。。這纔是老大！”，拱手曰:“各位兄弟，小生有禮!”"
case 10
ljjh_userinto="在江湖的大道上,%%開著一輛大奔（唉，真是舒服，氣死你！），腰間武器多多，來到了"&Application("ljjh_chatroomname")&",“呵。。這纔是真正的老大！”，拱手曰:“各位兄弟，小生有禮!”"
end select
if info(4)>0 then
ljjh_userinto="在江湖的大道上,%%開著一輛金裝的寶馬轎車（唉，全身包裝，金鏤衣），腰間掛著各式會員卡、車上裝滿了金銀珠寶，來到了"&Application("ljjh_chatroomname")&"，呵，做會員就是好啊,見到了江湖的兄弟們，蹬著眼睛叫到“別惹我，小心我給你喫會員卡”"
end if
if info(6)="掌門" then
ljjh_userinto="在江湖的大道上,%%開著一輛國產破踏板（黑煙直冒），腰間一把五4，身後跟著兩個彪形大漢，來到了"&Application("ljjh_chatroomname")&"，“哼、偶是掌門偶怕誰呀”。見到了本派眾位弟子,曰:“小的們，有什麼幫派事務趕快說來!我還急著去逛怡紅院呢”"
end if
if info(0)=Application("ljjh_admin") then
ljjh_userinto="在靈劍江湖總站飛機場上，站長%%乘坐『靈劍戰機』從空中徐徐降落，護駕的是20餘個飛行中隊的近百架飛機。恨天走下靈劍戰機向朋友們揮揮手，然後喊到：“我來了,大家有事請快說!”"
end if
if info(5)="通緝犯" then
	ljjh_userinto="警報，江湖通緝犯%%已經潛入我們內部，正想作一些不可告人的勾當，各全大俠小心為好   ”"
end if

if info(13)<>"無" then
sd(199)="<font color=black>【公告】</font><font color=008800>" & Replace(ljjh_userinto,"%%","<img src='../ico/"& info(10) &"-2.gif' width='84' height='84'><a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('["& info(9) &"]') target=f2>" & info(0) &"</a>["&info(5)&"]["&info(6)&"]<font color=red><b>id:"& info(9) &"</b></font>帶著自己的孩子["&info(13)&"] <img src='../ico/"& info(10) &"-2.gif' width='16' height='16'>") & "</font><font class=t>(" & t & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
else
sd(199)="<font color=black>【公告】</font><font color=008800>" & Replace(ljjh_userinto,"%%","<img src='../ico/"& info(10) &"-2.gif' width='84' height='84'><a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('["& info(9) &"]') target=f2>" & info(0) &"</a>["&info(5)&"]["&info(6)&"]<font color=red><b>id:"& info(9) &"</b></font>") & "</font><font class=t>(" & t & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
end if

sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
'此end if與上面的站長進入公告對應
end if
end if
if info(1)="女" then
if InStr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," " & LCase(pe) & " ")<>0 then
mess="老婆大人來了，還不快去請安！"
call messages(mess,pe)
end if
end if
sub messages(mess,pe)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=1
sd(194)="消息"
sd(195)=pe
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【系統】"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
sub messages2(mess)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=1
sd(194)="消息"
sd(195)=info(0)
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【系統提示】"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
chatroomname=Application("ljjh_chatroomname"&session("nowinroom"))
for i=0 to chatroomnum
ljchat=split(ljjh_roominfo(i),"|")
roomlist=roomlist&"<img src='room.gif'><a href=\"&chr(34)&"goroom1.asp?roomsn="&i&"&chatroomname="&ljchat(0)&"\"&chr(34)&" target=\"&chr(34)&"d\"&chr(34)&" title=\"&chr(34)&"  "&ljchat(3)&"\"&chr(34)&">"&ljchat(0)&"</a>"
next
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html;charset=big5">
<style>BODY{background-image: URL(bj.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}\n");</style>
<script Language="Javascript">
//self.resizeBy(screen.height,screen.width);
if(window!=window.top)
{
window.alert("請使用ie瀏覽器使用本系統！");
top.location.href="../exit.asp"
}
if(window.name!="xajh2004")
{ var i=1;
while (i<=50)
{
window.alert("你想作什麼呀，黑我？這裡是不行的，去別處玩去吧！哈！慢慢點50次！！");
i=i+1;
}
top.location.href="READONLY/BOMB.HTM"
}
var crm='<%=chatroomname%>';
var myn='<%=info(0)%>';
var cs=<%=info(2)%>;
var lst=0;
var tbclu=true;
var mdcls=true;
var listfaces=false;
var showhao=false;
var showpy="<%=info(11)%>";
var showmp="<%=info(5)%>";
var showsex="<%=info(1)%>";
var showseek="";
var chatcls = 0 ;
var clsok=0 ;
var bc = 1 ;
var showtype = 0 ;
headhigh=20;
var badword = new Array("射精","奸","去死","吃屎","你媽","你娘","日你","尻","操你","﹞") ;
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/‵ㄐ﹞ㄒㄓㄔ＃˙＆＊ㄩ※§〞ㄙㄗ﹛ㄘ〞ㄚ?ㄜˊ﹜ㄞ﹝ㄛˋ▲◎" ;

document.write("<title>"+crm+"</title>");
function write(cls){var fsize,lheight;
if(cls==1){fsize=this.f2.document.af.fs.value;lheight=this.f2.document.af.lh.value;}else{fsize='9';lheight='120';}
this.f1.document.open();
this.f1.document.writeln("<html><head><title>對話區</title><meta http-equiv=Content-Type content=\"text/html; charset=big5\">");
this.f1.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:125%}.t{color:FF00FF;font-size:11pt;}body{font-family:\"新細明體\";font-size:11pt}A{text-decoration:none;color:FFFF00}A:Hover{text-decoration:underline;color:FFFF00}A:visited{color:FFFF00}BODY{background-image: URL('1c.gif');background-position:'right top';background-repeat: no-repeat;background-attachment: fixed;overflow-x:hidden;}</style></head>");
this.f1.document.writeln("<\script language=javascript>");
this.f1.document.writeln("var autoScroll=1;var scTimer;speedscroll=0");
this.f1.document.writeln("function scrollit(){if(!parent.f2.document.af.as.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}");
this.f1.document.writeln("function speed(){if(!parent.f2.document.af.sp.checked){speedscroll=0;}else{speedscroll=1;}}");
this.f1.document.writeln("function autoscrollnow(){if (speedscroll==0 && autoScroll==1){this.scroll(0,65000);parent.f0.scroll(0,65000);setTimeout('autoscrollnow()',200);}else{var f1t=document.body.scrollTop;var f0t=parent.f0.document.body.scrollTop;var f1h=document.body.scrollHeight;var f0h=parent.f0.document.body.scrollHeight;if(autoScroll==1){if(f1t<=f1h||f0t<f0h){f1t+=15;f0t+=15;this.scroll(0,f1t);parent.f0.window.scroll(0,f0t);scTimer=setTimeout(\'autoscrollnow()\',10);}}}}<\/script>");
this.f1.document.writeln("<body oncontextmenu=self.event.returnValue=false background='../1.jpg' bgproperties='fixed' bgcolor=000000 text=DFF0C0>" );
this.f1.document.writeln("<span class=l><font color=red>【刷新畫面】</font>熱烈歡迎<font color=red> "+myn+" </font>來到《"+crm+"》</span><br><\script language=javascript>autoscrollnow();<\/script>");
this.f1.document.writeln("<%=roomlist%><br><\script language=javascript>autoscrollnow();<\/script>");
this.f0.document.open();this.f0.document.writeln("<html><head><title>分屏顯示</title><meta http-equiv=Content-Type content=\"text/html; charset=big5\">");this.f0.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:11pt;}body{font-family:\"新細明體\";font-size:11pt;}A{text-decoration:none;color:FFFF00}A:Hover{text-decoration:underline;color:FFFF00}A:visited{color:FFFF00}body{font-family:\"新細明體\";font-size:11pt;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL('c2.gif');background-position:'right top';background-repeat: no-repeat;background-attachment: fixed;overflow-x:hidden;}</style></head>");this.f0.document.writeln("<body oncontextmenu=self.event.returnValue=false background='../2.jpg' bgproperties='fixed' bgcolor=000000 text=DFF0C0>");this.f0.document.writeln("<span class=l><font color=red>【刷新畫面】</font>熱烈歡迎<font color=red> "+myn+" </font>來到《"+crm+"》加入會員請點右下面的會員！</span><br>");this.t.location.href="t.asp";
}
function sh(n1id,n2id,mm,n1,n2,y1,y2,bq,ll){
chatcls=chatcls+1;
if (chatcls>400 && clsok==0)
{if(confirm("你的屏幕顯示發言總行數已經超過400行,考慮到您的電腦性能問題,系統自動清屏，加快聊天速度,是否清屏？點擊“確定”清屏，點擊“取消”以後不再出現此提示")){chatcls=0;parent.qp();}else{clsok=1;}}
var faces;
if (this.f2.document.af.soundtx.checked != true){ll=ll.replace(".wav","");ll=ll.replace(".mid","")}
if (listfaces==true){
faces="<img src='../ico/"+ this.f2.document.af.mdtx.value +"-2.gif' width='20' height='20'>"
}
else
{
faces="";
}
var show;
if (n2=="mingdan" & this.f2.document.af.listname.checked==true){n2="大家";if(parent.m.location.href=="about:blank"){parent.m.location.href="f3.asp";}else{parent.m.location.reload();}}
if(n1id==-1){
if(mm==1){show="<span class=l>【私聊】"+ll+"</span><br>";}
else{show="<span class=l>"+ll+"</span><br>";}
}
else{
if(mm==1){show="<span class=l>【私聊】";}
else{show="<span class=l>";}
if(n1==myn){show=show+faces+"<font color="+y1+">你</font>"+bq;}
else{show=show+"<a href=javascript:parent.sw('["+n1+"]');parent.sws('["+n1id+"]') target=f2><font color="+y1+">"+n1+"</font></a>"+bq;}
if(n2==myn){show=show+"<font color=red> <font color="+y1+">"+n2+"</font> </font>";}
else{show=show+"[<font color="+y1+">"+n2+"</font>]";}
show=show+"說："+"<font color="+y2+">"+ll+"</font></span><br>";}
if ((n1==myn | n2==myn)&tbclu==false){show="<div  style='background:#AAE9A3;'>"+show+"</div>"}
if(tbclu==true & (myn==n1 | myn==n2)){this.f1.document.writeln(show);}else{this.f0.document.writeln(show);}
}
function qp(){this.f0.location.href="about:blank";
this.f1.location.href="about:blank";setTimeout('parent.write(1)',500);}
function sw(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towho.options[0].value=usna;this.f2.document.af.towho.options[0].text=usna;this.f2.document.af.sytemp.focus();return;}
function sws(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towhoid.value=usna;this.f2.document.af.sytemp.focus();return;}

function md1(){this.f3.document.writeln("<html><head><meta http-equiv='content-type' content='text/html; charset=big5'><title>在線用戶列表</title><style type='text/css'>");
this.f3.document.writeln("body{font-family:新細明體;font-size:11pt;}td{font-family:新細明體;font-size:11pt;line-height:125%;}A{color:white;text-decoration:none;}A:Hover{color:yellow;text-decoration:none;}A:Active {color:ffcc00}</style>");
this.f3.document.writeln("<\script language=\"JavaScript\"><\/script>");
this.f3.document.writeln("</head><body oncontextmenu=self.event.returnValue=false  bgcolor=#A4530B  background=bk.jpg bgproperties=\"fixed\">");
}

function md2(){
var nowroom='<%=session("nowinroom")%>';
if (nowroom == 0){
this.f3.document.writeln("<br><a href=javascript:parent.sw('[大家]');parent.sws('[0]')>^_^大家^_^</a><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[超級出氣包]');parent.sws('[0]')><s>超級出氣包</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[戰神]');parent.sws('[0]')><s>戰神</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[狂殺大盜]');parent.sws('[0]')><s>狂殺大盜</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[劈靂金刀]');parent.sws('[0]')><s>劈靂金刀</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
}
if (nowroom == 1){
this.f3.document.writeln("<br><a href=javascript:parent.sw('[大家]');parent.sws('[0]')>^_^大家^_^</a><br>");
for(var gw=180;gw<390;gw++){gw=gw+30;
this.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"級怪物]');parent.sws('[0]') title=\"============&#13&#10怪物兇猛 切勿亂打&#13&#10============\";><font size=3 color=#FFFF00><img src='../ico/"+gw+".gif'  border='0' width='36' height='36'><s>"+gw+"級怪物</s></font></a>&nbsp;<br>");}
for(var gw=490;gw<780;gw++){gw=gw+50;
this.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"級怪物]');parent.sws('[0]') title=\"============&#13&#10怪物兇猛 切勿亂打&#13&#10============\";><font size=3 color=#FFFF00><img src='../ico/"+gw+".gif'  border='0' width='36' height='36'><s>"+gw+"級怪物</s></font></a>&nbsp;<br>");}
for(var gw=801;gw<1100;gw++){gw=gw+49;
this.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"級怪物]');parent.sws('[0]') title=\"============&#13&#10怪物兇猛 切勿亂打&#13&#10============\";><font size=3 color=#FFFF00><img src='../ico/"+gw+".gif'  border='0' width='36' height='36'><s>"+gw+"級怪物</s></font></a>&nbsp;<br>");}
this.f3.document.writeln("<a href=# onClick=window.open('hqg.asp','hqg','width=500,height=400');><br><img src='pic/h0.gif'  border='0'><font color=yellow><font size=3>洪七公</font></font></a>&nbsp;<br>");}
}
function tma(onlineok){
var ss;
var faces;
var myself;
var friend;
var colorok;
var hycolor;
ljjh_zt=onlineok.split("|")
if (showpy.search(ljjh_zt[2]) != -1){friend="<img src='../jhimg/friend.gif' width='12' height='12'>";}
else
{friend=""}
if (ljjh_zt[2]==myn){myself="<img src='../jhimg/self.gif' width='12' height='12'>";}
else
{myself="";
}
if (listfaces==true){
faces="<img src='../ico/"+ljjh_zt[8]+"-2.gif' width='"+headhigh+"' height='"+headhigh+"'>"+friend+myself;
}
else
{
faces="";
}
if(ljjh_zt[3]=="男"&&ljjh_zt[0]==0){hycolor="#DEECFD";faces="<img src='boy.gif'>";}
if(ljjh_zt[3]=="男"&&ljjh_zt[0]>0){hycolor="#20AFE0";faces="<img src='boy.gif'>";}
if(ljjh_zt[3]=="女"&&ljjh_zt[0]==0){hycolor="#FFD0CF";faces="<img src='girl.gif'>";}
if(ljjh_zt[3]=="女"&&ljjh_zt[0]>0){hycolor="#F04F00";faces="<img src='girl.gif'>";}
if(ljjh_zt[1]=="官府"){colorok="#7FCF50";faces="<img src='room.gif'>";}else colorok="#FFFF00";
if(ljjh_zt[5]=="掌門"){colorok="#FFC01F";faces="<img src='room.gif'>";}
if(ljjh_zt[1]=="站長"){colorok="8FD060";faces="<img src='room.gif'>";}
ss=faces+"<a href=\"javascript:parent.sw(\'[" + ljjh_zt[2] + "]\');\"";
ss=ss+"><font size=2 color="+hycolor+">"+ljjh_zt[2]+"</font></a>"+" <font size=2 color="+colorok+">"+ljjh_zt[1]+"</font><br>";
//對於僅顯示好友的處理
this.f3.document.writeln(ss);
}
function mdsm(){
if (listfaces==false){
this.f3.document.writeln("<br></table>");
;}
}
function md3(){this.f3.document.writeln("</td></tr></table></body></html>");this.f3.document.close();}
function fc(){rn();setTimeout('parent.write()',0);}
function rn(){this.f2.document.af.username.value=myn;}
function clsay(){if(cs<6){setTimeout("this.f2.document.af.sytemp.value=''",0);}}
function IsBadWord(m)
{var tmp = "" ;
for(var i=0;i<m.length; i++)
{for(var j=0;j<badstr.length;j++)
if(m.charAt(i) == badstr.charAt(j)) break;
if(j==badstr.length) tmp += m.charAt(i) ;}
for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;
return false;}
function Warning()
{this.f2.document.af.sytemp.value='';
if(bc > 3) top.location.href="autokick.asp";else{bc++;alert("請不要在聊天室內說髒話,否則踢出！！！");}}
function checksays()
{if(this.f2.document.af.addvalues.checked)
{alert("您目前打開了自動泡點功能，不能發言！")
return false;}
if(this.f2.document.af.sytemp.value=="/炸彈$"){window.open("maninfosearch.asp?id="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(this.f2.document.af.sytemp.value=="/查看物品$"){window.open("towupin.asp?toname="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(this.f2.document.af.sytemp.value=="/信息$"){window.open("../yamen/mt.asp?action="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(IsBadWord(this.f2.document.af.sytemp.value)){Warning();return false;}
this.f2.document.af.sy.value='';if(this.f2.document.af.sytemp.value!='')
{
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=this.f2.document.af.sytemp.value;
this.f2.document.af.command.options[0].selected=true;
this.f2.document.af.command2.options[0].selected=true;
this.f2.document.af.command3.options[0].selected=true;
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.value='';ty=new Date();var nh=ty.getHours();var nm=ty.getMinutes();var ns=ty.getSeconds();var ct=(nh*3600)+(nm*60)+ns;if((ct-lst)>2.5){lst=ct;}else{this.f2.af.sytemp.value=this.f2.af.oldsays.value;
this.f2.af.oldsays.value='';return false;}this.f2.addOne(this.f2.document.af.sy.value);
this.f2.startnosay();
this.f2.document.af.subsay.disabled=1;setTimeout('this.f2.document.af.subsay.disabled=0',3000);
return true;}if ((this.f2.document.af.zt.options[this.f2.document.af.zt.selectedIndex].value=='0')||(this.f2.document.af.zt.options[this.f2.document.af.zt.selectedIndex].value=='')){alert('請輸入發言或選擇動作！');
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.select();return false;}}
function tbclutch(){
if(this.f2.document.af.tbclutch.value=='全屏')
{this.f2.document.af.tbclutch.value='垂直';
this.msgfrm.rows = "*";
this.msgfrm.cols = "*";
tbclu=false;}
else
{
if(this.f2.document.af.tbclutch.value=='垂直')
{this.f2.document.af.tbclutch.value='水平';
this.msgfrm.cols = "*,*";
this.msgfrm.rows = "*";
tbclu=true;}
else
{
this.f2.document.af.tbclutch.value='全屏';
this.msgfrm.cols = "*";
this.msgfrm.rows = "*,*";
tbclu=true;}}
this.f2.document.af.sytemp.focus();}
function face(){
if(this.f2.document.af.listface.checked==true){
parent.m.location.reload();
listfaces=true;
}
else{
parent.m.location.reload();
listfaces=false;
}
this.f2.document.af.sytemp.focus();
}
self.onerror=null;
var nullframe = '<HTML><BODY BGCOLOR=#000000 text=#ffffff><center><H3 color=yellow><font color=yellow>歡迎來到靈劍江湖</font></h3><br><br>程序正在從服務器讀取資料, 請稍候 ......</center></BODY></HTML>';
</script>
</head>
<frameset name=mainfrm cols="*" rows="*,0" border=0>
<frameset cols="*,154" name=tbgn1 rows="*" border="0" framespacing="0" frameborder="NO">
<frameset rows="*,0,70,0" cols="*">
<frameset name=msgfrm rows="*,*" cols="*">
<frame src="javascript:parent.nullframe" name="f0" scrolling="AUTO" marginheight="3" marginwidth="5" style="border-top:1px #9A87DE dashed;border-left:1px #9A87DE dashed">
<frame src="javascript:parent.nullframe" name="f1" scrolling="AUTO" marginheight="3" marginwidth="5" style="border-top:1px #9A87DE dashed;border-left:1px #9A87DE dashed">
</frameset>
<frame src="about:blank" name="t" marginwidth="5" marginheight="5" noresize scrolling="NO">
<frame src="f2.asp" name="f2" scrolling="NO" marginwidth="3" marginheight="8" noresize  style="border-top:1px #9A87DE dashed;border-left:1px #9A87DE dashed">
<frame src="option.asp" name="d" scrolling="NO">
</frameset>
<frameset rows="0,510,0,150" cols="*" name=tbymd>
<frame src="f3.asp" name="m" noresize scrolling="NO">
<frame src="about:blank" marginwidth="5" marginheight="5" name="f3">
<frame src="about:blank" marginwidth="5" marginheight="5" name="ps">
<frame src="f4.asp" scrolling="NO" noresize name="f4" marginwidth="3" marginheight="3">
</frameset>
</frameset>
<frame name="optfrm" src="about:blank" marginheight=0 marginwidth=0 scrolling=no noresize>
</frameset><noframes></noframes></html>
