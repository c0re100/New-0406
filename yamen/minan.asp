<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
PageSize = 15
rs.open "delete * from �H�R where �ɶ�<now()-3",conn,3,3
seekname=Trim(Request.Form("name"))
if seekname<>"" then
	rs.open "Select * From �H�R where ����='"&seekname&"' or ����='"&seekname&"'Order by �ɶ� DESC",conn,3,3
else
	rs.open "Select * From �H�R Order by �ɶ� DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>����R��</title>
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body leftmargin="0" topmargin="0" background="../001/tanpopo7.jpg">
<p><br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <img src="../001/treehouse1.gif" width="399" height="228" usemap="#MapMap" border="0"> 
  <script language=javascript> 
     function Click(){
     alert('�F�C�����w��z!!!');
     window.event.returnValue=false;
     }
     document.oncontextmenu=Click;
     </script>
  <map name="MapMap">
    <area shape="rect" coords="56,24,330,180" a href ="listminan.asp"  alt="�d�ݦ��̦�H" title="�d�ݦ��̦�H" >
  </map>
  <map name="Map">
    <area shape="rect" coords="44,201,102,252" a href ="listlao.asp" onClick="window.open('listlao.asp','yamen','scrollbars=no,resizable=no,top=25,left=100,width=570,height=400')" alt="�d�ݦb��H��" title="�d�ݦb��H��" >
  </map>
</p>
</body> 
<html><script language="JavaScript">                                                                  </script></html>