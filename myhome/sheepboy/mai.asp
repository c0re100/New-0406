<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("shop.asa")
Conn.Open connstr
shopname=int(Request.QueryString("shopname"))
sql="SELECT * FROM 商店 where id="&shopname
rs.open sql,conn,1,1
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=484"
shang=rs("商店名")
user=rs("店主")

rs.close
rs.Open "SELECT * FROM 商店物品  where 擁有者="&shopname&" and 數量>0",conn
if rs.EOF or rs.BOF then
rs.close
set rs=nothing
conn.close
set conn=nothing
if user=info(0) then
%><head>
<script language="vbscript">
alert("此店暫時沒有物品出售！")
location.href = "modifyshop.asp"
</script>
<%else
%>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>
<head>
<script language="vbscript">
alert("此店暫時沒有物品出售！")
location.href = "javascript:window.close()"
</script>
<%
Response.End 
end if
end if
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【小道消息】["&shang&"]老板["&user&"]對著顧客["&info(0)&"]說：歡迎呀，多買些東西啊........</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<title>商店</title>
</head>

<body bgcolor="#0066CC" LINK="#99FFCC" leftmargin="0" topmargin="0" >
 
<map name="MapMap2Map"> 
  <area shape="rect" coords="296,3,343,28" href="modifyshop.asp" target="_blank" alt="老板管理" title="老板管理" >
  </map>
  
<map name="MapMap2"> 
  <area shape="rect" coords="128,39,175,64" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=法器','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="法器類" title="法器類">
    
  <area shape="rect" coords="78,39,125,64" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=法寶','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="法寶類" title="法寶類">
    
  <area shape="rect" coords="26,39,74,65" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=鮮花','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="鮮花類" title="鮮花類">
  </map>
    
<map name="MapMap3"> 
  <area shape="rect" coords="76,32,123,57" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=槍械','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="槍類" title="槍類">
    
  <area shape="rect" coords="26,31,73,56" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=裝飾','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="裝飾類" title="裝飾類">
    
  <area shape="rect" coords="27,61,74,86" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=禮物','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="禮物類" title="禮物類">
    
  <area shape="rect" coords="128,32,175,57" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=彈藥','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="子彈類" title="子彈類">
    
  <area shape="rect" coords="77,61,124,86" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=礦石','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="礦石類" title="礦石類">
    
  <area shape="rect" coords="127,4,174,29" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=毒藥','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="毒藥類" title="毒藥類">
    
  <area shape="rect" coords="76,3,123,28" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=藥品','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="藥品類" title="藥品類">
    
  <area shape="rect" coords="128,62,175,87" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=頭盔','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="頭盔類" title="頭盔類">
    
  <area shape="rect" coords="27,3,74,28" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=暗器','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="暗器類" title="暗器類">
  </map>
    
<map name="Map"> 
  <area shape="rect" coords="465,-1,561,49" href="#" onClick="window.close()" alt="關閉窗口" title="關閉窗口">
    
  <area shape="rect" coords="78,2,125,27" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=雙腳','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="腳部裝備類" title="腳部裝備類">
    
  <area shape="rect" coords="15,29,95,54" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=左手','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="左手裝備" title="左手裝備">
    
  <area shape="rect" coords="27,1,74,26" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=盔甲','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="盔甲類" title="盔甲類">
    
  <area shape="rect" coords="112,29,192,54" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=右手','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="右手裝備" title="右手裝備">
  </map>
  <p align="center"> 
<div id="Layer1" style="position:absolute; width:105px; height:38px; z-index:1; left:11px; top:0px"><img src="logo.gif" width="88" height="31"></div>
<p align="center"><FONT SIZE="3" COLOR="#000000"><img src="img/001.jpg" width="640" height="60" border="0"><img src="img/002.jpg" width="640" height="52" border="0"><img src="img/003.jpg" width="640" height="52" usemap="#MapMap2Map" border="0"><img src="img/004.jpg" width="640" height="70" border="0"><img src="img/005.jpg" width="640" height="68" border="0" usemap="#MapMap2"><img src="img/006.jpg" width="640" height="90" usemap="#MapMap3" border="0"><img src="img/007.jpg" width="640" height="88" usemap="#Map" border="0"></FONT>
