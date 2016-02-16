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
rs.open "select id from 用戶 where id="&info(9),conn
dim page
page=request.querystring("page")
PageSize = 15
myname=request.cookies("Jname")
myname=info(0)
set buys=server.createobject("adodb.recordset")
buys.open "delete * from  殺手 where 時間<now()-10",conn,3,3
buys.open "SELECT * FROM 殺手 order by 時間 DESC",conn,3,3
buys.PageSize = PageSize
pgnum=buys.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then buys.AbsolutePage=page
%>
<html>

<head>
<title>雇傭殺手</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#660000" background="../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="100%" border="1" bgcolor="#990000">
  <tr>
    <td><font color="#FFFFFF">完成任務後，請來此聲明(<font color="#FFFF00">3分鐘內</font>)，否則遲了，就沒有錢可領了！ 
      三級以上纔可以雇傭殺手殺人!官府人員不可雇傭! 只有殺人組織（門派為“刺客”）的人可以被雇傭！ 所以各位大俠還是快快用盡用力吧!超過七日被雇傭都沒有將人殺死，就會賠出,同樣超過七日沒有人應征殺人，也會賠出。。。（切記一個人隻能雇傭一次，否則多次雇傭後，有一個七天內雇傭者沒將人殺死，將會將所有同樣的名字的記錄都刪掉）注意：雇傭的殺手殺不殺人都不會有所損失的</font></td>
  </tr>
</table>
<table width="100%" align="center" cellspacing="2" border="0" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7"> 
    <td align="left"><font size="2">[共<font color="#990000"><b><%=buys.pagecount%></b></font>頁][<font
color="#990000"><%=musers()%></font>人雇傭殺手] [<a href=shashoulist.asp?page=<%=buys.pagecount-1%>><font color="#0000FF">上一頁</font></a>][第<%=page%>頁][<a href=shashoulist.asp?page= <%=buys.pagecount+1%>><font color="#0000FF">下一頁</font></a>]</font></td>
  </tr>
  <tr> 
    <td> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%"
bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="12%"><font size="2">請求殺人者</font></td>
          <td width="13%"><font size="2">即將被殺者</font></td>
          <td width="10%"><font size="2">被雇擁者</font></td>
          <td width="13%"><font size="2">殺人傭金</font></td>
          <td width="18%"><font size="2">殺人說明</font></td>
          <td width="21%"><font size="2">請求殺人時間</font></td>
          <td width="13%"><font size="2">是否成功</font></td>
        </tr>
        <%
count=0
do while not buys.eof and count<buys.PageSize
%>
        <tr bgcolor="#f7f7f7"> 
          <td width="12%"><font size="2"><%=buys("姓名1")%></font></td>
          <td width="13%"><font size="2"><%=buys("被殺者")%></font></td>
          <td width="10%"><font size="2"><%=buys("姓名2")%></font></td>
          <td width="13%"><font size="2"><%=buys("殺人傭金")%></font></td>
          <td width="18%"><font size="2"><%=buys("說明")%></font></td>
          <td width="21%"><font size="2"><%=buys("時間")%> </font> 
          <td width="13%"><font size="2"> 
            
            <a
href="shashouling.asp?name= mmTranslatedValueHiliteColor="HILITECOLOR=%22Dyn%20Untranslated%20Color%22"<%=buys("被殺者")%>"><font color="#FF0000">領錢</font></a> 
            
            <%if info(2)=10 then %>
            <a
href="shashoulingok.asp?name= mmTranslatedValueHiliteColor="HILITECOLOR=%22Dyn%20Untranslated%20Color%22"<%=buys("姓名1")%>"><font color="#FF0000">超過時間刪掉</font></a> 
            <%end if%>
            </font></td>
        </tr>
        <%buys.movenext%>
        <%count=count+1%>
        <%loop
buys.close
rs.close
%>
      </table>
    </td>
  </tr>
</table>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As 姓名1 from 殺手")
musers=tmprs("姓名1")
set tmprs=nothing
if isnull(musers) then musers=0
end function
rs.open "select 等級 from 用戶 where id="&info(9),conn
denji=rs("等級")
rs.close
conn.close
set rs=nothing	
set conn=nothing
%><p>&nbsp;</p>
<p align="center"> 
  <%if denji>3 then%>
  <a href="gushashou.htm"><font size="2">我要雇傭殺手</font></a> <a href="javascript:this.location.reload()"><font size="2" >刷新</a> 
  <%end if%>
</p>
</body>

</html>