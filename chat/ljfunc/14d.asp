<%
function attackd(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('拚命對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
info(4)=kkman
if kkman=0 then kkman=1
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
Response.End 
end if
if instr(fn1,"&")=0 or right(fn1,1)="&" then
Response.Write "<script language=JavaScript>{alert('操作錯誤，格式如下：[武器名&子彈數量]');}</script>"
Response.End 
end if
zt=split(fn1,"&")
abc=left(trim(zt(1)),1)
if abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('操作錯誤，子彈數量請使用數字！');}</script>"
Response.End 
end if
aszcc=zt(1)
winname=zt(0)
winname=Replace(winname," ","")
aszcc=Replace(aszcc," ","")
if abs(aszcc)>50 then
	Response.Write "<script language=JavaScript>{alert('子彈不能發射超過50發！');}</script>"
Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select top 1 姓名1,姓名2 FROM 單挑 ",conn
xx1=rs("姓名1")
xx2=rs("姓名2")
rs.close
rs.open "select 攻擊,武功,性別,內力,社團,殺人數,保護,子彈,攜帶武器,武器威力,暴豆時間,等級 from 用戶 where 姓名='"&info(0)&"'",conn
eeee=rs("攜帶武器")
if rs("社團")="出家" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經出家了，不可以再殺人了！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
dim dwArr

if rs("攜帶武器")<>winname then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你沒有配帶"&winname&"這種武器！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("子彈")<50 then
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script language=JavaScript>{alert('哈哈，你沒有這麼多子彈呀？！你的子彈要補充到50發才可以射擊');}</script>"
Response.End
exit function
end if
if rs("殺人數")>=30 then
	conn.execute "update 用戶 set 保護=false where 姓名='"&info(0)&"'"
	Response.Write "<script language=JavaScript>{alert('你作惡太多，超過江湖殺人上限30不能再發招了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("保護")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('你閉關修練中..不能射擊！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if xx1=info(0) or xx2=info(0) then
diantiaook=1
else
diantiaook=0
end if
okok=rs("武器威力")
if okok=0 then okok=10
wgnl=okok*aszcc
xb=rs("性別")
if rs("殺人數")>=30 then
conn.execute "update 用戶 set 保護=false where 姓名='"&info(0)&"'"
Response.Write "<script language=JavaScript>{alert('你作惡太多，超過江湖殺人上限"& Application("sjjh_killman") &"不能再發招了！');}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
sj=DateDiff("n",rs("暴豆時間"),now())
xinxi=""
baodoudj=1
if sj<=60 then
xinxi="<font color=red>威力爆滿</font>"
baodoudj=3
end if
rs.close

'對方的判斷
rs.open "select 姓名,狀態,性別,防御,武功,會員等級,保護,社團 from 用戶 where 姓名='"&toname&"'",conn
toname=rs("姓名")
yxjfyto=rs("防御")
yxjwgto=rs("武功")
jhhy=rs("會員等級")
xb1=rs("性別")
menpai1=rs("社團")
if rs("社團")="出家" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('[" & to1 & "]已經出家，不能攻擊！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("保護")=True then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('["&toname&"]保護中..不能攻擊！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if rs("狀態")="死" then
Response.Write "<script language=JavaScript>{alert('不行呀;"& toname &"已經掛了你們誰給他復活呀');}</script>"
Response.End
end if
rs.close
randomize timer
kaa=int(rnd*999)+1
killer=int(wgnl*2)
killer=killer*kkman
killer=killer+kaa
if killer>=1000000 then
randomize timer
killer=int(rnd*999999)+1
end if
if killer<=10 then
randomize timer
killer=int(rnd*99)+1
end if
'對方的判斷
randomize timer
conn.execute "update 用戶 set 體力=體力-"  & killer & " where 姓名='"&toname&"'"
rs.open "select 體力 from 用戶 where 姓名='"&toname&"'",conn
if rs("體力")>-100 then
randomize
aaa=int(rnd*6)
select case aaa
case 1
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc*3&">用"& winname &"<img src='pic/5.gif'>射擊" & toname & ",發射了<font color=red><b>"&aszcc&"</b></font>發子彈,使得"& toname &"體力下降:<font color=red>-"& killer &"</font>點."
case 2
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc*3&">用"& winname &"<img src='pic/5.gif'>射擊" & toname & ",發射了<font color=red><b>"&aszcc&"</b></font>發子彈,使得"& toname &"體力下降:<font color=red>-"& killer &"</font>點."
case 3
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc*3&">用"& winname &"<img src='pic/5.gif'>射擊" & toname & ",發射了<font color=red><b>"&aszcc&"</b></font>發子彈,使得"& toname &"體力下降:<font color=red>-"& killer &"</font>點."
case else
attackd=toname&"對<img src='pic/dz44.gif'>"&info(0)&"大叫道:哈哈,只是給我搔癢癢,再來呀"
end select
rs.close
set rs=nothing
conn.close
set conn=nothing
else
conn.execute "update 用戶 set 狀態='死' where 姓名='"&toname&"'"
conn.execute "update 用戶 set 銀兩=銀兩+20000,體力=體力+10000,內力=內力+10000 where id="&info(9)
rs.close
attackd=xinxi & info(0) & "<bgsound src=wav/dcc.wav loop="&aszcc&">用"& winname &"射擊" & toname & "只聽一聲慘叫，"& toname &"躺在地上說道：啊<img src='pic/dz30.gif'>我死得好慘啊，然後眼睛一閉與世長辭了.....<bgsound src=mp3/013.mp3 loop=2>"
set rs=nothing
conn.close
set conn=nothing
end if
end function
%>