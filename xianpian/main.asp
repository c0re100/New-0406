<%@ Language=VBScript %>
<% option explicit%>
<!--#include file="adovbs.asp"-->
<!-- #include file="conn.asp" -->
<%session("mynameup")=""%>
<html><head>
<script language='javascript' SRC="tip.js"></script>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<STYLE type=text/css>BODY {
FONT-SIZE: 12px; FONT-FAMILY: "新細明體","Arial Narrow","Times New Roman"
}
P {
FONT-SIZE: 12px; COLOR: #82668a; FONT-FAMILY: "新細明體","Arial Narrow","Times New Roman"
}
TD {
FONT-SIZE: 12px; COLOR: #82668a; FONT-FAMILY: "新細明體","Arial Narrow","Times New Roman"
}
A:link {
FONT-SIZE: 12px; COLOR: #978de9; FONT-FAMILY: 新細明體; TEXT-DECORATION: none
}
A:visited {
FONT-SIZE: 12px; COLOR: #978de9; TEXT-DECORATION: none
}
P {
FONT-SIZE: 12px; FONT-FAMILY: 新細明體
}
DIV {
FONT-SIZE: 12px; FONT-FAMILY: 新細明體
}
A:hover {
COLOR: red
}
</STYLE>
<script LANGUAGE="JavaScript">
bg = new Array(33);
bg[0] = '../riji/000.gif'; bg[1] = '../riji/001.gif'; bg[2] = '../riji/002.gif'; bg[3] = '../riji/003.gif'; bg[4] = '../riji/004.gif'; bg[5] = '../riji/005.gif'; bg[6] = '../riji/006.gif'; bg[7] = '../riji/007.gif'; bg[8] = '../riji/008.gif'; bg[9] = '../riji/009.gif'; bg[10] = '../riji/010.gif'; bg[11] = '../riji/011.gif'; bg[12] = '../riji/012.gif'; bg[13] = '../riji/013.gif'; bg[14] = '../riji/014.gif'; bg[15] = '../riji/015.gif'; bg[16] = '../riji/016.gif'; bg[17] = '../riji/017.gif'; bg[18] = '../riji/018.gif'; bg[19] = '../riji/019.gif'; bg[20] = '../riji/010.gif'; bg[21] = '../riji/021.gif'; bg[22] = '../riji/022.gif'; bg[23] = '../riji/023.gif'; bg[24] = '../riji/024.gif'; bg[25] = '../riji/025.gif'; bg[26] = '../riji/026.gif'; bg[27] = '../riji/027.gif'; bg[28] = '../riji/028.gif'; bg[29] = '../riji/029.gif'; bg[30] = '../riji/030.gif'; bg[31] = '../riji/031.gif'; bg[32] = '../riji/032.gif';
index = Math.floor(Math.random() * bg.length); document.write("<BODY BACKGROUND="+bg[index]+">"); </script>
<title>靈劍交友</title>
<meta http-equiv="Page-Enter" content="revealTrans(Duration=1.0,Transition=23)">
</head>
<body topmargin="0" leftmargin="0" background="../ljimage/11.jpg" bgcolor="#FFFFFF">
<div align="center">
<center>


    <table border="0" width="100%" cellspacing="0" cellpadding="2" background="../ljimage/11.jpg">
      <tr>
        <td width="17%" align="center">  
          <div ID="tipBox" style="width: 109; height: 32"></div>
        </td>
        <td width="83%" align="center"><img border="0" src="02.gif" width="528" height="85" align="left"></td>
</tr>
</table>

</center>
</div>

<div align="center"><center>
<table border="1" width="727" cellspacing="0" cellpadding="3" style="font-family: 新細明體; font-size: 9pt; color: #000000" bordercolor="#988CD0">
<tr>
<td bgcolor="#F8EBD1" align="center" width="145"><a href="main.asp?action=search">⊙-所有大俠</a></td>
<td bgcolor="#F8EBD1" align="center" width="145"><a href="add.asp">⊙-大俠加入</a></td>
<td bgcolor="#F8EBD1" align="center" width="145"><a href="userre.asp">⊙-修改資料</a></td>
<td bgcolor="#F8EBD1" align="center" width="145"><a href="sf.asp?Keys=Login">⊙-超級管理</a></td>
<td bgcolor="#F8EBD1" align="center" width="145"><a href="admsearch.asp?keys=admsearch">⊙-高級搜索</a></td>
</tr>
</table>
</center></div>

<%

dim list,strSQL
const MaxPage=5
dim CurrentPage
dim totalPut
dim TotalPages
dim i,j,action
dim Aff,m
Aff = Chr(34)

if request("action")="search" then
if not isempty(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
'----------------STart--
'Response.Write request("keyword")    '反饋搜索數據
'----------------ENd----
select case request("kind")
case "name"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  name    LIKE '%"&request("keyword")&"%'"
case "sex"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  sex     LIKE '%"&request("keyword")&"%'"
case "age"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  age     LIKE '%"&request("keyword")&"%'"
case "xingzuo"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  xingzuo LIKE '%"&request("keyword")&"%'"
case "icqnum"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  icqnum  LIKE '%"&request("keyword")&"%'"
case "likes"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  likes   LIKE '%"&request("keyword")&"%'"
case "liuyan"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  doc     LIKE '%"&request("keyword")&"%'"
case "address"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  address LIKE '%"&request("keyword")&"%'"
case "mail"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  mail    LIKE '%"&request("keyword")&"%'"
case "url"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  url     LIKE '%"&request("keyword")&"%'"
case "mode"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  mode    LIKE '%"&request("keyword")&"%'"

case "gsm"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  gsm    LIKE '%"&request("keyword")&"%'"
case "pc"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  pc    LIKE '%"&request("keyword")&"%'"
case "danwei"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & "  danwei    LIKE '%"&request("keyword")&"%'"

case "alls"
strSQL="SELECT * FROM List WHERE "
strSQL=strSQL & " name       LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or sex     LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or age     LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or xingzuo LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or icqnum  LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or likes   LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or doc     LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or address LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or mail    LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or url     LIKE '%"&request("keyword")&"%'"
strSQL=strSQL & " or mode    LIKE '%"&request("keyword")&"%'"
case else
strSQL="SELECT * FROM List "
end select
If request.querystring("m")<>"" Then
Select case request.Querystring("m")
case "counter"
strSQL =strSQL & " Order by counter DESC"
m      = "counter"
case "name"
strSQL =strSQL & " Order by  name   DESC"
m      = "name"
case "sex"
strSQL =strSQL & " Order by  sex    DESC"
m      = "sex"
case "age"
strSQL =strSQL & " Order by  age    DESC"
m      = "age"
case "mode"
strSQL =strSQL & " Order by  mode   DESC"
m      = "mode"
case "ip"
strSQL =strSQL & " Order by  ip     DESC"
m      = "ip"
case "times"
strSQL =strSQL & " Order by  times  DESC"
m      = "times"
End Select
Else
strSQL =strSQL & " Order by ID DESC"
m      = ""
End if
'-------STart--
'   Response.Write strSQL     '<<@@@----------顯示SQL語法是否正確！
'-------ENd----
set list=server.createobject("adodb.recordset")

list.open strSQL,conn,adOpenKeySet,adLockPessimistic
if list.eof and list.bof then
%>
<div align="center"><center>
<table border="1" width="95%" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: 新細明體; font-size: 9pt; color: #000000">
<tr><td width="100%">
<p align="center"><font color="#FF0000">
Sorry! 沒有你想找的資料！</font></p>
</td></tr></table></center></div>
<%
searchForm
else
totalPut=list.recordcount
if currentPage=1 then
showContent
showpages
searchForm
else
if (currentPage-1)*MaxPage<totalPut then
list.move  (currentPage-1)*MaxPage
dim bookmark
bookmark=list.bookmark
showContent
showpages
searchForm
else
currentPage=1
showContent
searchForm
end if
end if
list.close
end if

set list=nothing

else
searchForm

sub searchForm()%>
<form method=post action=main.asp name=searchform >
<div align="center"><center><table border="1" width="95%" cellspacing="0" cellpadding="3" bordercolor="#988CD0" bgcolor="#E0E0E0" style="font-family: 新細明體; font-size: 9pt; color: #000000">
<tr><td width="100%"><p align="center">
<select name="kind" style="font-family: 新細明體; font-size: 9pt" size="1">
<option value="alls" selected>按字段搜索</option>
<option value="name">按姓名搜索</option>
<option value="mode">按班級搜索</option>
<option value="sex">按性別搜索</option>
<option value="age">按年齡搜索</option>
<option value="xingzuo">按星座搜索</option>
<option value="icqnum">按ICQ 搜索</option>
<option value="likes">按電話搜索</option>
<option value="gsm">按手機呼機</option>
<option value="liuyan">按留言搜索</option>
<option value="pc">按郵政編碼</option>
<option value="danwei">按單位名稱</option>
<option value="address">按聯繫地址</option>
<option value="mail">按電子郵件</option>
<option value="url">按主頁地址</option>
</select>
<input name="keyword" size="20"  style="font-family: 新細明體; font-size: 9pt">
<input name=action type=hidden value=search>
<input name="ok" style="font-family: 新細明體; font-size: 9pt;BACKGROUND-COLOR: #e7e7de; PADDING-TOP: 1pt" type="submit"
value="<%
if request("action")="search" then
Response.Write "新搜索"
else
Response.Write "搜 索"
end if
%>">
</p></td></tr>
</table></center></div></form>
<%
end sub

end if

sub showcontent
dim cal

cal=1
%>
<div align="center"><center>
    <table border="0" width="95%" cellspacing="0" cellpadding="3" style="font-family: 新細明體; font-size: 9pt; color: #000099" background="../ljimage/11.jpg">
      <tr> 
        <td align="center">序號</td>
        <td align="center"><a href="main.asp?action=search&m=name" title="↓">姓名</a></td>
        <td align="center"><a href="main.asp?action=search&m=mode" title="↓">門派</a></td>
        <td align="center"><a href="main.asp?action=search&m=sex" title="↓">性別</a></td>
        <td align="center"><a href="main.asp?action=search&m=age" title="↓">年齡</a></td>
        <td align="center"><a href="main.asp?action=search&m=ip" title="↓">大俠</a><a href="main.asp?action=search&m=ip" title="↓">照片</a></td>
        <td align="center"><a href="main.asp?action=search&m=times" title="↓">登記時間</a></td>
        <td align="center"><a href="main.asp?action=search&m=counter" title="↓">瀏覽次數</a></td>
      </tr>
      <%do while not (list.eof or err)%>
      <tr> <form method=post action=lookuser.asp name=lookuser >
        <input name=id type=hidden value=<%=list("id")%>>
        <td align="center"><%= i+1 +(currentPage-1)*MaxPage%></td>
        <td align="center">
          <%'用戶的名稱和ID號碼！%>
          <a href=lookuser.asp?ID=<%=list("ID")%> onMouseOver="this._tip = '<%=list("name")%>&nbsp;<font color = blue><%=list("sex")%></font>&nbsp;<%=list("age")%>歲<br>生日:<%=list("mons")%>-<%=list("days")%><br>星座:<%=list("xingzuo")%><br><font color = 8080FF>詳細信息,請</font><font color=red>點擊</font><font color = 8080FF>進入</font>'"><font class=pt9 color=#0000FF><%=list("name")%></font></a> 
        </td>
        <td align="center"><a href="main.asp?action=search&kind=mode&keyword=<%=list("mode")    %>"><%=list("mode")    %></td>
        <td align="center"><%=list("sex")    %></td>
        
        </td>
        <td align="center"> 
          <% if isnull(list("photo")) or isempty(list("photo")) then%>
          還沒有上傳相片 
          <%else %>
          <a href="lookuser.asp?ID=<%=list("ID")%>"><img src="showimg.asp?id=<%=List("name")%>" width="64" height="56"> 
          </a> 
          <%End If%>
        </td>
        <td align="center"><%=list("times")  %></td>
        <td align="center"><%=list("counter")%></td>
      </tr>
      <% i=i+1
if i>=MaxPage then exit do
list.movenext
loop
%>
    </table>
  </center></div></form>

<%end sub

sub showpages()
dim n
if (totalPut mod MaxPage)=0 then
n= totalPut \ MaxPage
else
n= totalPut \ MaxPage + 1
end if
if n=1 then exit sub
%>
<div align="center"><center>
<table border="1" width="95%" cellspacing="0" cellpadding="3" bordercolor="#988CD0" style="font-family: 新細明體; font-size: 9pt; color: #000000" background="../ljimage/11.jpg">
  <tr width=100%><td>現有<font color="#FF6600"><!--#include file="zongshu.asp" --></font>位同學加入,共<%=n%>頁：

<%	   dim k
for k=1 to n
if k=currentPage then
response.write "<h>" & Cstr(k)+" </h>"
else
If m<>"" Then
response.write "<a href='main.asp?action=search&m="+m+"&kind="+request("kind")+"&keyword="+request("keyword")+"&page="+cstr(k)+"'>"+Cstr(k)+"</a> "
Else
response.write "<a href='main.asp?action=search&kind="+request("kind")+"&keyword="+request("keyword")+"&page="+cstr(k)+"'><font size=3><b>"+Cstr(k)+"</b></font></a> "
End if
end if
next%>
</td><td align=right>當前為第<%=currentPage%>頁</td>
</tr>
</table><%end sub%>
<!--#include file="copy.asp" -->
</body></html>

