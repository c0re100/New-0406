<%	
'遊戲開始初始化
Sub GameStart(bet)
	
call CheckMoney(bet)
		
Session("n_Begin")=true '是否已開局
Session("n_Init")=true '是否初次進入本頁面
Session("n_Bet")=bet '所押錢數
				
redim userPokers(1,0)
redim antiPokers(1,0)
Session("n_UserPokers")=userPokers '用戶牌
Session("n_AntiPokers")=antiPokers '電腦牌
Session("n_PokerBack")=GetPokerBack() '本副牌背面圖案
Session("n_Result")=4 '結果狀態號
Session("n_TotalMoney")=Session("n_TotalMoney")-bet
	
End Sub
	
'遊戲結束處理
Sub GameOver()
	
Session("n_Result")=Result()	
Session("n_Begin")=false
		
dim value	
			
value=CCur(Session("n_Bet"))
			
select case CInt(Session("n_Result"))
case 1
value=value*2
case 2
value=0
case 3
value=value
case else
value=0
end select
		
Session("n_TotalMoney")=Session("n_TotalMoney")+value
Session("n_Bet")=0
	
End Sub
	
Sub GiveUserPoker()
		
dim userPokers
		
userPokers=Session("n_UserPokers")
if UBound(userPokers,2)<5 then
Session("n_UserPokers")=AddPoker(userPokers,"User")		
end if
			
End Sub
	
Sub GiveAntiPoker()
		
dim antiPokers
		
antiPokers=Session("n_AntiPokers")
		
if UBound(antiPokers,2)<5 then
Session("n_AntiPokers")=AddPoker(antiPokers,"Anti")
end if
		
End Sub
'得到牌背面圖片文件名
Function GetPokerBack()
		
dim rndNum
		
Randomize
rndNum=Int((7 * Rnd) + 1)
		
GetPokerBack="deck_" & rndNum & ".gif"
		
End Function
	
'顯示計算機的牌
Sub ShowAntiPokers()
		
dim pokers
dim outStr
		
pokers=Session("n_AntiPokers")
		
outStr="<table width='100%' border=0><tr><td>"
	
if UBound(pokers,2)>0 then
for i=1 to UBound(pokers,2)
if not Session("n_Begin") then
outStr=outStr & GejhimgURL(GetPokerFace(pokers(0,i),pokers(1,i))) & "&nbsp;&nbsp;"
else
outStr=outStr & GejhimgURL(Session("n_PokerBack")) & "&nbsp;&nbsp;"	
end if
next
else
outStr=outStr & "&nbsp;"
end if
	
outStr=outStr & "</td></tr></table>"
	
Response.Write outStr
	
End Sub
	
'得到牌面圖片文件名
Function GetPokerFace(byval pokerNum,byval pokerType)
			
select case CLng(pokerNum)
case 1
pokerNum="a"
case 11
pokerNum="j"
case 12
pokerNum="q"
case 13
pokerNum="k"
case else
pokerNum=pokerNum
end select
		
GetPokerFace=pokerType & "_" & pokerNum & ".gif"
		
End Function
	
'得到牌背面圖片文件名
Function GetRndPokerBack()
		
dim rndNum
dim pokerType
		
Randomize
rndNum = Int((4 * Rnd) + 1)
		
select case CLng(rndNum)
case 1
pokerType="h"
case 2
pokerType="d"
case 3
pokerType="s"
case 4
pokerType="c"
end select
		
GetRndPokerBack=pokerType
		
End Function
	
'顯示用戶牌
Sub ShowUserPokers()
	
dim pokers
dim outStr
		
pokers=Session("n_UserPokers")
	
outStr="<table width='100%' border=0><tr><td>"
		
if UBound(pokers,2)>0 then
for i=1 to UBound(pokers,2)
outStr=outStr & GejhimgURL(GetPokerFace(pokers(0,i),pokers(1,i))) & "&nbsp;&nbsp;"
next
else
outStr=outStr & "&nbsp;"
end if
	
outStr=outStr & "</td></tr></table>"
	
Response.Write outStr
		
End Sub
	
'得到所給牌所能達到的最大點數
Function GetMaxPoint(pokers)
		
dim point '牌的點數
dim aCount '牌中A的個數
		
point=0
aCount=0
GetMaxPoint=0
		
'算牌中A的個數
for i=1 to UBound(pokers,2)
select case pokers(0,i)
case 1
aCount=aCount+1
point=point+1
case 11
point=point+10
case 12
point=point+10
case 13
point=point+10
case else
point=point+pokers(0,i)
end select
next
	
'算當前最大點數
for i=1 to aCount
if (point+10)>21 then
exit for
else
point=point+10
end if
next
	
GetMaxPoint=point
		
End Function
	
'判斷電腦是否能加牌
Function CanAddPoker()
	
dim point
dim between
dim can
dim multi
		
pokers=Session("n_AntiPokers")
		
if UBound(pokers,2)<2 then
CanAddPoker=true
exit function
end if
		
point=GetMaxPoint(pokers)
		
between=21-point
can=between*4
		
multi=0
		
for i=1 to UBound(pokers,2)
if pokers(0,i)>=1 and pokers(0,i)<=between then
can=can-1
multi=multi+1
end if
next
		
if can/(54-multi)>=0.5 then
CanAddPoker=true
else
CanAddPoker=false
end if
		
End Function
	
	
'1=用戶贏
'2=計算機贏
'3=平手
Function Result()
		
dim yPt '用戶點數
dim cPt '計算機點數
dim aCount
dim cPokers
dim yPokeers
		
cPokers=Session("n_AntiPokers")
yPokers=Session("n_UserPokers")
		
cPt=GetMaxPoint(cPokers)
yPt=GetMaxPoint(yPokers)
	
if yPt>21 then
if cPt>21 then
Result=3
elseif yPt=cPt then
Result=3
else
Result=2
end if
elseif yPt=21 then
if cPt<>21 then
Result=1
else
Result=3
end if
else
if cPt>21 then
Result=1
else
if yPt>cPt then
Result=1
elseif yPt=cPt then
Result=3
else
Result=2
end if
end if
end if			
				
End Function
	
'對牌值的處理
Function Judge(byval poker)
	
if poke>10 then
Judge=10
else
Judge=poker
end if
		
End Function
	
'產生隨機牌
Function RandNum()

Randomize
RandNum = Int((13 * Rnd) + 1)

End Function
	
Function GetAntiNum(byval num)
GetAntiNum=num
End Function
	
'加入新牌
Function AddPoker(byval pokers,byval userType)
		
dim count
dim num
		
count=UBound(pokers,2)
count=count+1
redim preserve pokers(1,count)
		
num=RandNum()
		
select case UCase(userType)
case "USER"
num=GetUserNum(num)
case "ANTI"
num=GetAntiNum(num)
case else
num=GetUserNum(num)
end select
		
pokers(0,count)=num '數字
pokers(1,count)=GetRndPokerBack() '背面圖片
		
AddPoker=pokers
	
End Function
	
Function GetUserNum(byval num)
		
dim rnd
		
rnd=Int((5*Rnd)+1)
		
if rnd<>2 and rnd<>3 then
do until num<9
num=RandNum()
loop
end if
		
GetUserNum=num
		
End Function
	
'到到圖片路徑
Function GejhimgURL(byval picName)
		
if picName="" then
GejhimgURL=""
else
GejhimgURL="<img src='jhimg/" & picName & "' width=71 height=96>"
end if
		
End Function
	
'設置按鈕是否可選
Function EnableBtn(byval bool)
		
if bool then
EnableBtn=""
else
EnableBtn="disabled=true"
end if
			
End Function
	
Sub CheckMoney(byval money)
		
if IsNumeric(money) then
money=CCur(money)
if money>CCur(Session("n_TotalMoney")) then
call ErrorForm("你窮瘋了，不知道自己有多少錢呀！",1)
elseif money=0 then
call ErrorForm("你還沒押錢，賭什麼賭呀",1)
elseif money<0 then
call ErrorForm("想干什麼。。哼~",1)
elseif money<10 then
call ErrorForm("大哥，別這麼小氣嘛，最少也得押10兩銀子呀",1)
elseif money>1000 then
call ErrorForm("大哥，中國人厲害呀！我最多隻能接1000兩銀子的賭注",1)
end if
else
call ErrorForm("大哥，讓你押錢又沒讓你寫拼音",1)
end if
		
End Sub

Sub ErrorForm(str,path)
dim back

Response.Write "<p align='center'><font color='red'>" & str & "</font></p>"
		
if IsNumeric(path) then
back="<p align='center'><a href='javascript:{history.go(-" & path & ");}'>[返回]</a>"
else
back="<p align='center'><a href='" & path & "'>[返回]</a>"
end if

Response.Write back
Response.End

End Sub
%>