<%
function dpkfp(fn1,from1,to1)

if to1="大家" and fn1<>"認輸了" then
  response.write "<script Language='Javascript'>alert('請選擇說話對像，不能像大家發送這個命令!');</script>"
  response.end
end if
fn1=ucase(fn1)
chatroomsn=session("nowinroom")
fname=info(0)
if left(fn1,2)="公證" then
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then 
  response.write "<script Language='Javascript'>alert('錯誤的命令格式：\n正確示例：\n/發牌$ 公證|負\n/發牌$ 公證|勝!');</script>"
  response.end
end if
	fn_62=left(arr_fn1(1),1)

	if fn_62<>"負" and fn_62<>"勝" then 
  response.write "<script Language='Javascript'>alert('錯誤的命令格式：\n正確示例：\n/發牌$ 公證|負\n/發牌$ 公證|勝!');</script>"
  response.end
end if
	if info(2)<6 then 
  response.write "<script Language='Javascript'>alert('你不是管理員，不能進行牌局公證!');</script>"
  response.end
end if
	fname=to1
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/dpk.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 



sql="select * from dpk where ufrom='"& fname & "'"
Set Rs=conn.Execute(sql)

if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('["&fname&"]沒有參與[撲克]牌局!');</script>"
  response.end

end if
dpk=rs("pk")
dpkto=rs("uto")
xiazhu=rs("duz")
fpok=rs("fp")
oldoldpn=Cint(rs("oldpn"))
oldmaxpk=rs("maxpk")
logtime=rs("logtime")
iszai=rs("iszai")
rs.close
'dpkfp=cstr(oldoldpn)
'exit function

dim pk_geshi,old_pk_geshi
pk_geshi=1

if oldoldpn<10 then oldoldpn=90
old_pk_geshi=CINT(left(oldoldpn,1))
oldoldpn=CINT(mid(oldoldpn,2))

'------------------------新格式------------------------
dim xia_class,xia_fir,xia_sql
xia_fir=left(xiazhu,1)
xiazhu=mid(xiazhu,2)

select case xia_fir
	case "1"
		xia_class="金幣"
		xia_sql="金幣"
	case "2"
		xia_class="銀兩"
		xia_sql="銀兩"
	case "3"
		xia_class="法力"
		xia_sql="法力"
	case else
		conn.close
		set rs=nothing
		set conn=nothing
  response.write "<script Language='Javascript'>alert('非法操作!');</script>"
  response.end
end select
'------------------------新格式------------------------
d_sql="delete * from dpk where ufrom='" & fname & "' or uto='" & fname &"'"

if fn_62<>"" then

	if fn_62="負" then
		s_pk=dpkto
		f_pk=to1
	elseif fn_62="勝" then
		f_pk=dpkto
		s_pk=to1
	End if

	Set Rs=conn.Execute(d_sql)
	conn.close
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")


	'------------------------新格式------------------------
	sql="select " & xia_sql & " from 用戶 where 姓名='" & f_pk &"'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
		xiazhu=0
		sayxia="，【" & f_pk & "】已經改名換姓逃走了"
	else
		yin1=rs(""&xia_sql&"")
		if Clng(yin1)<Clng(xiazhu) then 
		xiazhu=yin1
		sayxia="，這家伙身上只有這麼多" & xia_class  & "了"
		end if
	end if
	rs.close
	
	conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where 姓名='"& s_pk &"'"
	conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where 姓名='"& f_pk &"'"
	conn.close

dpkfp="【官府公證】"&f_pk&"不遵守[撲克牌局]規定拒絕完成牌局不肯認輸<br>官府人員[ " & info(0) & " ]認為此人沒有賭品,裁定他輸給[ " & s_pk & " ]" & xiazhu & xia_class & sayxia 
elseif fn1<>"不要了" and fn1<>"認輸了" then
  if to1<>dpkto then
  response.write "<script Language='Javascript'>alert('你的對手不是[" & to1 & "]啊!');</script>"
  response.end
  end if
  if fpok then
  response.write "<script Language='Javascript'>alert('現在好像不是你發牌啊,等對方出牌了才是你發牌!');</script>"
  response.end
  end if
'olddpk22=split(dpk,"|")
'olddpk3=ubound(olddpk22)
dim qysall
dim qysn
dim pkx
dim isArrpk(15)
dim isArrpkn(15)

fn1=fn1 & "+"
fn1=replace(fn1,"++","+")
fn2=split(fn1,"+")
fn3=ubound(fn2)
for i=0 to fn3-1
  if instr(dpk,fn2(i) & "|")<>0 then
     oldpn=oldpn+1
     pkx=mid(fn2(i),2)
     if fn2(i)="大王" or fn2(i)="小王" then pkx=fn2(i)
     pkx=unconvpk(pkx)
     isArrpkn(pkx)=isArrpkn(pkx) + 1
     isArrpk(pkx)=isArrpk(pkx) & " " & showpk(fn2(i))
     qys=left(fn2(i),1)
     qysn=qysn+1
     if qysall<>qys then qysall=qysall & qys
  end if
  dpk=replace(dpk,fn2(i) & "|","")
next
maxpk=0
dim zaid,sang,yidu,pkyizha2,pkyizha3,liandui
zaid=0
sang=0
yidu=0
pkyizha2=0
pkyizha3=0
pkyizha3=1
liandui=2

for i=15 to 1 step -1
  if isArrpkn(i)=4 then
    zaid=zaid+4
    fpnow =fpnow & " 四個(炸)" & isArrpk(i)
    if maxpk=0 then 
	maxpk=i
	pk_geshi=4
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=3 then
    fpnow =fpnow & " 三個" & isArrpk(i)
    sang=sang+3
    if maxpk=0 then 
	maxpk=i
	pk_geshi=3
    end if
  end if
next

for i=15 to 1 step -1
  if isArrpkn(i)=2 then
     fpnow =fpnow & " 一對" & isArrpk(i)
     yidu=yidu+2
     if isArrpkn(i-1)=2 then liandui=liandui+2
     if maxpk=0 then 
	maxpk=i
	pk_geshi=2
     end if
  end if
next
for i=15 to 1 step -1
  if isArrpkn(i)=1 then
    pkyizha2=pkyizha2+1
    fpnow =fpnow & " 一張" & isArrpk(i)
    if isArrpkn(i-1)=1 then pkyizha3=pkyizha3+1
    if maxpk=0 then 
	maxpk=i
	pk_geshi=1
    end if
  end if
next
dpk22=split(dpk,"|")
dpk3=ubound(dpk22)
if len(qysall)=1 and qysn>1 then 
qysall="<b><font color=red>[清一色]</font></b>"
if maxpk=0 then 
	pk_geshi=pk_geshi+5
end if
else
qysall=""
end if

dim Errfp
Errfp=false
dim ErrStr
ErrStr=""
dim show_old_pk
show_old_pk=oldmaxpk

if oldmaxpk=11 then show_old_pk="J"
if oldmaxpk=12 then show_old_pk="Q"
if oldmaxpk=13 then show_old_pk="K"
if oldmaxpk=14 then show_old_pk="A"
if oldmaxpk=15 then show_old_pk="2"
if oldmaxpk=16 then show_old_pk="小王"
if oldmaxpk=17 then show_old_pk="大王"

        IF old_pk_geshi=2 then
		if oldoldpn<3 then
			e_Str="一對"
		else
			e_Str="連對"
		end if
	elseif old_pk_geshi=3 then
		e_Str="三個"
	elseif old_pk_geshi=4 then
		e_Str="炸彈"
	elseif old_pk_geshi=1 then
		if oldoldpn>4 then
			e_Str="一條龍最大的一張牌："
		else
			e_Str="一張" 
		end if
	end if
	if (old_pk_geshi-5>0 and oldoldpn<>0) then e_Str=e_Str & "清一色" 

if zaid<>4 and oldoldpn<>0 then
        
	if old_pk_geshi<>pk_geshi then
  response.write "<script Language='Javascript'>alert('["& to1 & "剛才出的是[" & e_Str & " " & show_old_pk & "]啊!');</script>"
  response.end
	elseif oldpn<>oldoldpn then
		'erralt(oldpn)
  response.write "<script Language='Javascript'>alert('["& to1 & "剛才出的是[" & oldoldpn & "]張啊!');</script>"
  response.end
	end if

	if ErrStr<>"" then
  response.write "<script Language='Javascript'>alert('如果你沒有這種牌，請在[洗牌]桌按[讓 牌]!');</script>"
  response.end
	end if
end if


if pkyizha3=oldpn and oldpn>4 then 
qysall=qysall & "<b><font color=red>[一條龍]</font></b>"
IsyitiaoL=true
else
IsyitiaoL=false
end if
if zaid=0 and yidu=0 and sang=0 and IsyitiaoL=false then
isgood=false
else
isgood=true
end if


if (yidu>0 and pkyizha2>0) then Errfp=true 
if (zaid=0 and (yidu>2 and yidu<>liandui) ) then Errfp=true '兩對以上不是連對
if (sang<>0 and yidu=0 and pkyizha2>2) then Errfp=true '有三個有一對還有一張以上零牌

if (sang<>0 and oldpn<>5 and oldpn<>3) then Errfp=true '三個大於五張 
if (zaid<>0 and oldpn>7) then Errfp=true '炸彈大於7張

if (oldpn=3 and yidu<>0) then Errfp=true '三張牌中有一對

if (pkyizha2>2 and sang<>0) then Errfp=true '如果兩張不是一對
if (isgood=false and oldpn>1) then Errfp=true '沒有炸沒有對沒有三個不是一條龍，而且不是一張牌
if Errfp=true then
  response.write "<script Language='Javascript'>alert('你想發的是沒有規律的零牌，這麼大的人打撲克也不會？看看幫助先！!');</script>"
  response.end
end if
if maxpk=14 then maxpk=16
if maxpk=15 then maxpk=17
'新的14.15必須在原來的1415後面
if maxpk=1 then maxpk=14
if maxpk=2 then maxpk=15

'response.write maxpk
'response.end
if maxpk<=Cint(oldmaxpk) then
  if (zaid>0 and iszai) or zaid=0 then
  response.write "<script Language='Javascript'>alert('["& to1 & "剛才出的牌[" & e_Str & " " & show_old_pk & "]比你大！如果你沒有牌比剛才的牌大，請在[洗牌]桌按[讓 牌]!');</script>"
  response.end
   end if
end if

if dpk3>0 then
sql="update dpk set pk='" & dpk & "',fp=true,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",logtime='" & now() & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=" & pk_geshi & oldpn & ",maxpk=" & maxpk & ",iszai=" & iszai & " where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
dpkfp=info(0)&"想了想，拿出[" & oldpn & "]張牌瀟灑的甩上了牌桌，原來是......" & ismore & "<br> " & qysall & fpnow & " , 還有" & dpk3 & "張牌"
%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
elseif oldpn=18 then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
conn.close
  response.write "<script Language='Javascript'>alert('你想作弊啊(不遵守遊戲規則想拋光所有的牌)？現在要罰你，請[" & dpkto & "]出牌!');</script>"
  response.end
else

Set Rs=conn.Execute(d_sql)
conn.close
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'------------------------新格式------------------------
sql="select " & xia_sql & " from 用戶 where 姓名='" & dpkto &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if yin1<0 then yin1=0
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia=",這家伙身上太窮了，只有這麼多" & xia_class  & "了"
end if

conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where 姓名='"& dpkto &"'"
conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "+"& xiazhu&" where 姓名='"& info(0)&"'" 
conn.close
dpkfp=info(0)&"哈哈大笑，把所有的牌發了出來......" & ismore & "<br> " &fpnow&",沒有牌了,從"&dpkto&"那贏到" & xiazhu & xia_class & sayxia
'------------------------新格式------------------------

end if

elseif fn1="不要了" then
sql="update dpk set fp=true,oldpn=90,maxpk=0 where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dpk set fp=false,oldpn=90,maxpk=0,iszai=false where ufrom='"& dpkto & "'"
Set Rs=conn.Execute(sql)
dpkfp=info(0)&"放棄打這張牌，請[" & dpkto & "]出牌"
conn.close
elseif fn1="認輸了" then

Set Rs=conn.Execute(d_sql)
conn.close
	Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")


'------------------------新格式------------------------
sql="select " & xia_sql & " from 用戶 where 姓名='" & info(0) &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia="，這家伙身上只有這麼多" & xia_class  & "了"
end if

conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where 姓名='"& dpkto &"'"
conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where 姓名='"& info(0) &"'"
conn.close
dpkfp=info(0)&"認輸了放棄打這局牌,輸給[" & dpkto & "]" & xiazhu & xia_class & sayxia
'------------------------新格式------------------------

end if
set conn=nothing
set rs=nothing
end function

function showpk(pk)
dim wpk
wpk=replace(pk,"紅","<img src=duju/dpk/1.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"黑","<img src=duju/dpk/2.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"梅","<img src=duju/dpk/3.GIF border=0><font color=#000000 size=2>")
wpk=replace(wpk,"方","<img src=duju/dpk/4.GIF border=0><font color=#FF0000 size=2>")
wpk=replace(wpk,"小王","<img src=duju/dpk/xw.gif border=0>")
wpk=replace(wpk,"大王","<img src=duju/dpk/dw.gif border=0>")
showpk=wpk & "</font>"
end function
function unconvpk(cpk)
dim pk2
pk2=replace(cpk,"A","1")
pk2=replace(pk2,"J","11")
pk2=replace(pk2,"Q","12")
pk2=replace(pk2,"K","13")
pk2=replace(pk2,"小王","14")
pk2=replace(pk2,"大王","15")
unconvpk=pk2
end function
%>