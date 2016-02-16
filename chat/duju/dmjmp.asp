<%
function dmjGetmj(from1,to1)
if to1="大家" and fn1<>"認輸了" then
  response.write "<script Language='Javascript'>alert('請選擇說話對像，不能像大家發送這個命令!');</script>"
  response.end
end if

chatroomsn=session("nowinroom")

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/dmj.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select * from dmj where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)

if rs.eof and rs.bof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('您沒有參加打麻將，怎麼就想到摸牌了!');</script>"
  response.end
else
	isMy=rs("isMy")
	isFp=rs("isFp")
	isGet=rs("isGet")
	Mymj=rs("Mymj")
	dmjto=rs("uto")
	xiazhu=rs("duz")
	mjID=rs("mjID")
	logtime=rs("logtime")
	rs.close

if not isMy then
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('現在好像不是你摸牌啊,等對方出牌了才是你摸牌!');</script>"
  response.end
end if
if isFp and (not isGet) then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你現在應當打牌(出牌),在[洗牌]桌點擊麻將使用(/打麻將 XXX) 命令!');</script>"
  response.end
end if
sql="select 麻將 from mjInfo where id="& mjID 
Set Rs=conn.Execute(sql)
Getmjs=rs("麻將")
rs.close

if len(Getmjs)>39 then
Getmjs2=right(Getmjs,3)
Mymj=Mymj & Getmjs2
Getmjs=left(Getmjs,len(Getmjs)-3)
sql="update mjInfo set 麻將='" & Getmjs & "' where id="& mjID 
Set Rs=conn.Execute(sql)
sql="update dmj set Mymj='" & Mymj & "',isGet=false,isFp=true,logtime='" & now() & "',mjmp='" & left(Getmjs2,2) & "' where ufrom='"& info(0) & "'"
Set Rs=conn.Execute(sql)
'last=strmj2(left(Getmjs2,2))

%>
<script language="JavaScript">parent.f3.location.href="duju/dmj-xp.asp";alert("你摸到<%=strmj2(left(Getmjs2,2))%>")</script>
<%
dmjGetmj=info(0)&",在牌桌摸了一張牌，小心的看了看，瞪了瞪眼，不知在想什麼......"
else
sql="delete * from dmj where ufrom='" & info(0) & "' or ufrom='" & dmjto &"'"
Set Rs=conn.Execute(sql)
sql="delete * from mjInfo where id=" & mjID
Set Rs=conn.Execute(sql)
dmjGetmj=info(0)&",只有13張牌了，不能再摸牌，你們打成和局，沒有勝負。"
end if
end if
End Function
%>