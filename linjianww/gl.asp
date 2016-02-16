<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim tmprs
tmprs=conn.execute("Select count(*) As 數量 from 用戶 where times<3 and CDate(lasttime)<now()-10")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
user1=musers
tmprs=conn.execute("Select count(*) As 數量 from 用戶 where allvalue<=5 and CDate(lasttime)<now()-10")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
user2=musers
tmprs=conn.execute("Select count(*) As 數量 from 用戶 where CDate(lasttime)<now()-30")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers

%><title>靈劍江湖數據庫維護程序</title><style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
<body background="0064.jpg">
<div align="center">
<p>主用戶數據庫：</p>
<table width="500" border="1" bordercolor="#000000" cellspacing="0" cellpadding="1">
<tr>
<td width="85%">10天之前,僅登陸2次的非會員有：<font color="#0000FF"><b><%=user1%>個</b></font></td>
<td width="15%">
<div align="center"><a href="gl1.asp">刪除</a></div>
</td>
</tr>
<tr>
<td width="85%">10天之前泡點數allvalue小於5的非會員有：<b><font color="#0000FF"><%=user2%>個</font></b></td>
<td width="15%">
<div align="center"><a href="gl2.asp">刪除</a></div>
</td>
</tr>
<tr>
<td width="85%">
        <p>30天從未登陸的非會員有：<font color="#0000FF"><b><%=user3%>個</b></font></p>
      </td>
<td width="15%">
<div align="center"><a href="gl3.asp">刪除</a></div>
</td>
</tr>
</table>
  <br>
  <p align="center"><BR>
    在這裡你可以查看到當前數據庫的使用情況，<br>
    選擇刪除將可以把這一些用戶刪除掉，刪除之後選擇數據庫壓縮操作！<br>
    在我的江湖上我就用這一些方法成功把一個11MB的數據庫壓縮成4MB大小！<br>
    在此操作前建議備份原始文件，如果因為程序操作造成的後果我們不負任何責任！<br>
    對於:記錄數據庫、BBS數據庫如果特別大的可以選擇用原始文件覆蓋安即可，<br>
    建議對這些文件備份<BR>
</div> 