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
sql="SELECT * FROM �ө� where id="&shopname
rs.open sql,conn,1,1
if rs.EOF or rs.BOF then Response.Redirect "../error.asp?id=484"
shang=rs("�ө��W")
user=rs("���D")

rs.close
rs.Open "SELECT * FROM �ө����~  where �֦���="&shopname&" and �ƶq>0",conn
if rs.EOF or rs.BOF then
rs.close
set rs=nothing
conn.close
set conn=nothing
if user=info(0) then
%><head>
<script language="vbscript">
alert("�����ȮɨS�����~�X��I")
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
alert("�����ȮɨS�����~�X��I")
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�p�D�����j["&shang&"]�ѪO["&user&"]����U��["&info(0)&"]���G�w��r�A�h�R�ǪF���........</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<title>�ө�</title>
</head>

<body bgcolor="#0066CC" LINK="#99FFCC" leftmargin="0" topmargin="0" >
 
<map name="MapMap2Map"> 
  <area shape="rect" coords="296,3,343,28" href="modifyshop.asp" target="_blank" alt="�ѪO�޲z" title="�ѪO�޲z" >
  </map>
  
<map name="MapMap2"> 
  <area shape="rect" coords="128,39,175,64" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�k��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�k����" title="�k����">
    
  <area shape="rect" coords="78,39,125,64" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�k�_','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�k�_��" title="�k�_��">
    
  <area shape="rect" coords="26,39,74,65" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�A��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�A����" title="�A����">
  </map>
    
<map name="MapMap3"> 
  <area shape="rect" coords="76,32,123,57" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�j��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�j��" title="�j��">
    
  <area shape="rect" coords="26,31,73,56" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�˹�','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�˹���" title="�˹���">
    
  <area shape="rect" coords="27,61,74,86" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=§��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="§����" title="§����">
    
  <area shape="rect" coords="128,32,175,57" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�u��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�l�u��" title="�l�u��">
    
  <area shape="rect" coords="77,61,124,86" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�q��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�q����" title="�q����">
    
  <area shape="rect" coords="127,4,174,29" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�r��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�r����" title="�r����">
    
  <area shape="rect" coords="76,3,123,28" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�ī~','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�ī~��" title="�ī~��">
    
  <area shape="rect" coords="128,62,175,87" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�Y��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�Y����" title="�Y����">
    
  <area shape="rect" coords="27,3,74,28" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�t��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�t����" title="�t����">
  </map>
    
<map name="Map"> 
  <area shape="rect" coords="465,-1,561,49" href="#" onClick="window.close()" alt="�������f" title="�������f">
    
  <area shape="rect" coords="78,2,125,27" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=���}','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�}���˳���" title="�}���˳���">
    
  <area shape="rect" coords="15,29,95,54" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=����','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="����˳�" title="����˳�">
    
  <area shape="rect" coords="27,1,74,26" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=����','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="������" title="������">
    
  <area shape="rect" coords="112,29,192,54" href="#" onClick="window.open('maiok.asp?shang=<%=shang%>&lei=�k��','fb','scrollbars=yes,resizable=no,width=500,height=400')" alt="�k��˳�" title="�k��˳�">
  </map>
  <p align="center"> 
<div id="Layer1" style="position:absolute; width:105px; height:38px; z-index:1; left:11px; top:0px"><img src="logo.gif" width="88" height="31"></div>
<p align="center"><FONT SIZE="3" COLOR="#000000"><img src="img/001.jpg" width="640" height="60" border="0"><img src="img/002.jpg" width="640" height="52" border="0"><img src="img/003.jpg" width="640" height="52" usemap="#MapMap2Map" border="0"><img src="img/004.jpg" width="640" height="70" border="0"><img src="img/005.jpg" width="640" height="68" border="0" usemap="#MapMap2"><img src="img/006.jpg" width="640" height="90" usemap="#MapMap3" border="0"><img src="img/007.jpg" width="640" height="88" usemap="#Map" border="0"></FONT>
