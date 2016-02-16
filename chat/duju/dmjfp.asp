<%
'搓麻將
function dmjfp(fn1,from1,to1)
if to1="大家" and fn1<>"認輸了" then 
  response.write "<script Language='Javascript'>alert('請選擇說話對像，不能像大家發送這個命令!');</script>"
  response.end
end if

chatroomsn=session("nowinroom")

fname=info(0)
if left(fn1,2)="公證" then
	arr_fn1=split(fn1,"|")
	if ubound(arr_fn1)<>1 then 
  response.write "<script Language='Javascript'>alert('錯誤的命令格式：\n正確示例：\n/發牌 公證|負\n/發牌 公證|勝!');</script>"
  response.end
end if

	fn_62=left(arr_fn1(1),1)

	if fn_62<>"負" and fn_62<>"勝" then 
  response.write "<script Language='Javascript'>alert('錯誤的命令格式：\n正確示例：\n/發牌 公證|負\n/發牌 公證|勝!');</script>"
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
db="duju/dmj.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& fname & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('"& fname & "沒有參與[麻將]牌局!');</script>"
  response.end
end if
isMy=rs("isMy")
isFp=rs("isFp")
isGet=rs("isGet")
Mymj=rs("Mymj")
dmjto=rs("uto")
xiazhu=rs("duz")
logtime=rs("logtime")
mjID=rs("mjID")
rs.close

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
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('非法操作!');</script>"
  response.end
end select
'------------------------新格式------------------------
d_sql="delete from dmj where ufrom='" & fname & "' or uto='" & fname &"'"
d_sql2="delete from mjInfo where id=" & mjID

if fn_62<>"" then
	if fn_62="負" then
		s_pk=dmjto
		f_pk=to1
	elseif fn_62="勝" then
		f_pk=dmjto
		s_pk=to1
	End if

	Set Rs=conn.Execute(d_sql)
	Set Rs=conn.Execute(d_sql2)
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
		if clng(yin1)<clng(xiazhu) then 
			xiazhu=yin1
			sayxia="，這家伙身上只有這麼多" & xia_class  & "了"
		end if
	end if
	rs.close
	
	conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where 姓名='"& s_pk &"'"
	conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "-"&xiazhu&" where 姓名='"& f_pk &"'"
	conn.close
	dmjfp="【官府公證】" & f_pk &",不遵守[麻將牌局]規定拒絕完成牌局不肯認輸<br> 官府人員[ " & info(0) & " ]認為此人沒有賭品,裁定他輸給[ " & s_pk & " ]" & xiazhu & xia_class & sayxia
elseif fn1<>"認輸了" then
if to1<>dmjto then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你的對手不是[" & to1 & "]啊!');</script>"
  response.end
end if
if not isMy then
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('現在好像不是你打牌啊,等對方出牌了才是你打牌!');</script>"
  response.end
end if
if isGet and (not isFp) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你應當先摸牌，然後再打牌,在[洗牌]桌點擊[摸 牌]使用(/摸牌)命令!');</script>"
  response.end

	%>
	<script language="JavaScript">parent.f2.document.forms[0].sytemp.value="/摸牌";</script>
	<%
	conn.close
	set rs=nothing
	set conn=nothing
	exit function
end if
if instr(Mymj,fn1 & "|")<>0 then
	Mymj=replace(Mymj,fn1 & "|","",1,1,1)
else
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你有沒有搞錯啊，你沒有這張牌啊!');</script>"
  response.end
end if
sql="update dmj set Mymj='" & Mymj & "',oldmj='" & fn1 & "',isGet=true,isFp=false,isMy=false,logtime='" & now() & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
sql="update dmj set isGet=true,isFp=false,isMy=true where ufrom='"& dmjto & "'"
Set Rs=conn.Execute(sql)
conn.close
dmjfp=info(0)&"拿出一張牌用手指頭搓了搓,嘿......擲上了牌桌，原來是 " & showMj(fn1) &"!"
%>
<script language="JavaScript">parent.f3.location.reload();parent.f2.document.forms[0].sytemp.value="/摸牌";</script>
<%
elseif fn1="認輸了" then
Set Rs=conn.Execute(d_sql)
Set Rs=conn.Execute(d_sql2)
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'------------------------新格式------------------------
sql="select " & xia_sql & " from 用戶 where 姓名='" & info(0) &"'"
set rs=conn.execute(sql)
yin1=rs(""&xia_sql&"")
if yin1<0 then yin1=0
if clng(yin1)<clng(xiazhu) then 
xiazhu=yin1
sayxia="，這家伙身上太窮了，只有這麼多" & xia_class  & "了"
end if

conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "+"&xiazhu&" where 姓名='"& dmjto &"'"
conn.execute "update 用戶 set " & xia_sql & "=" & xia_sql & "-"& xiazhu&" where 姓名='"& info(0) &"'"
conn.close
dmjfp=info(0)&",認輸不打了，輸給[" & dmjto & "]" & xiazhu  & xia_class & sayxia
'------------------------新格式------------------------
end if
set conn=nothing
set rs=nothing
end function
function showMj(Mj)
showMj="<input type=image border=0 src=""duju/mj/" & intMjp(Mj) & ".gif"" onclick=""javascript:parent.f3.location.href='duju/dmj-xp.asp';"" >"
end function
function strmj2(Mj)
dim Mj2
Mj2=replace(Mj,"1","一")
Mj2=replace(Mj2,"2","二")
Mj2=replace(Mj2,"3","三")
Mj2=replace(Mj2,"4","四")
Mj2=replace(Mj2,"5","五")
Mj2=replace(Mj2,"6","六")
Mj2=replace(Mj2,"7","七")
Mj2=replace(Mj2,"8","八")
strmj2=replace(Mj2,"9","九")
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL=left(cmj,1)
if isNumeric(mjL) then 
mj2=right(cmj,1) & mjL
mj2=replace(mj2,"索","1")
mj2=replace(mj2,"筒","2")
mj2=replace(mj2,"萬","3")
else
mj2=replace(mj2,"東風","10")
mj2=replace(mj2,"南風","20")
mj2=replace(mj2,"西風","30")
mj2=replace(mj2,"北風","40")
mj2=replace(mj2,"紅中","41")
mj2=replace(mj2,"白板","42")
mj2=replace(mj2,"發財","43")
end if
intMjp=mj2
end function
%>
