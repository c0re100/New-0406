<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
 
<!--#include file="../data2.asp"-->

<%
if info(0)="" then
%>
<script language="Vbscript">
msgbox"非法入侵！",0,"警告"
history.back
</script>
<%
Response.End
end if%>
<%
'rs.open "select * from diarymessage where clicknumber>counted and user='"&info(0)&"'",conn,1,3

'if not rs.bof then
'tempcount=0

'do while not rs.eof

'tempcount=tempcount+(rs("clicknumber")-rs("counted"))
'rs.update "counted",rs("clicknumber")
'rs.movenext
'loop
'rs.close
'rs.open "select * from bardian where user='"&info(0)&"'",conn,1,3
'if not rs.bof then
'rs.update "magic",tempcount*0.01
'end if
'rs.close
'else

'rs.close
'end if
%>
<%
pagecount=9
if request.querystring("page")="" or (not isnumeric(request.querystring("page")))then
dqpage=1
else
dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if
end if

rs.open "select * from diarymessage where user='"&info(0)&"' order by id desc ",conn,1,1

temp1=rs.recordcount/pagecount
if temp1-int(rs.recordcount/pagecount)<>0 then
totalpage=int(temp1+1)
else
totalpage=temp1
end if
if dqpage>totalpage then
dqpage=1
end if

if dqpage<>1 then
temp2=(dqpage-1)*pagecount
rs.move temp2

end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>啟航社區–我的留言板</title>
<link rel="StyleSheet" href="../../city/new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

<div align="center"><center>
</center></div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">□-您當前的位置--俠客寓所---我的日記本&nbsp;</td>
<td width="10"></td>
<td class=p1 width="440">□-當前時間：<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">□-日記撰寫區：</td>
<td width="10"></td>
<td class=p1 width="440">□-日記閱覽區：</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td width="326">
<form method="POST" action="mydiary_post.asp">

<table border="0" width="324" cellspacing="1" cellpadding="0">
<tr>
<td class=p3 width="314">
<p align="right"><select size="1" name="weather" style="font-family: Tahoma; font-size: 12px">
<option value="晴朗" selected>晴朗</option>
<option value="陰天">陰天</option>
<option value="多雲">多雲</option>
<option value="雨">雨</option>
<option value="雪">雪</option>
<option value="風">風</option>
<option value="霧">霧</option>
</select><select size="1" name="mood" style="font-family: Tahoma; font-size: 12px">
<option value="普通">普通</option>
<option value="甜蜜">甜蜜</option>
<option value="快樂">快樂</option>
<option value="高興">高興</option>
<option value="驚喜">驚喜</option>
<option value="悲傷">悲傷</option>
<option value="失落">失落</option>
<option value="無聊">無聊</option>
<option value="痛苦">痛苦</option>
<option value="寂寞">寂寞</option>
<option value="感動">感動</option>
<option value="興奮">興奮</option>
<option value="終生難忘">終生難忘</option>
<option value="失戀">失戀</option>
<option value="氣憤">氣憤</option>
<option value="無助">無助</option>
<option value="傷心">傷心</option>
</select></td>
<td class=p3 width="10">
</td>
</tr>
<center>
<tr>
<td class=p3 width="324" colspan="2"><textarea rows="14" name="detail" cols="59" style="font-family: Tahoma; font-size: 12px"></textarea></td>
</tr>
<tr>
<td class=p3 width="324" colspan="2">這篇日記是否開放日記：<input type="radio" value="yes" name="open">是<input type="radio" value="no" checked name="open">否
<input type="submit" value="提交" name="B2" style="font-family: Tahoma; font-size: 12px">
<input type="reset" value="重寫" name="B1" style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</form>
</center>
</td>
<td width="10"></td>
<td width="440" valign="top">
<table border="0" width="436" cellspacing="1" cellpadding="0"><FORM METHOD="get" ACTION="index.asp">
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="186" colspan="4">當前<%=dqpage%>頁： 總<%=totalpage%>頁：</td>
<td class=p2 width="164"></td>
<td class=p2 width="70"></td>
</tr></FORM>
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="46">

<%if dqpage>1 then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="1">
<INPUT TYPE="submit" value="首頁" style="font-family: Tahoma; font-size: 12px">
</FORM>
<%else%>
<INPUT TYPE="submit" value="首頁" style="font-family: Tahoma; font-size: 12px">
<%end if%>

</td>
<td class=p2 width="46">
<%if dqpage>1 then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="<%=dqpage-1%>">
<INPUT TYPE="submit" value="上頁" style="font-family: Tahoma; font-size: 12px">
</FORM>

<%else%>
<INPUT TYPE="submit" value="上頁" style="font-family: Tahoma; font-size: 12px">
<%end if%>
</td>
<td class=p2 width="46">
<%if dqpage<totalpage then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="<%=dqpage+1%>">
<INPUT TYPE="submit" value="下頁" style="font-family: Tahoma; font-size: 12px">
</FORM>


<%else%>
<INPUT TYPE="submit" value="下頁" style="font-family: Tahoma; font-size: 12px">
<%end if%>
</td>
<td class=p2 width="48">
<%if dqpage<totalpage then%>
<FORM METHOD=get ACTION="index.asp">
<INPUT TYPE="hidden" name="page" value="<%=totalpage%>">
<INPUT TYPE="submit" value="尾頁" style="font-family: Tahoma; font-size: 12px">
</FORM>


<%else%>
<INPUT TYPE="submit" value="尾頁" style="font-family: Tahoma; font-size: 12px">
<%end if%>
</td>
<FORM METHOD=get ACTION="index.asp">

<td class=p2 width="234" colspan="2">轉到第：<INPUT TYPE="text" NAME="page" size="4" maxlength=4 style="font-family: Tahoma; font-size: 12px">頁<input type="submit" value="確定" name="B1" style="font-family: Tahoma; font-size: 12px"></td></FORM>
</tr>
</table>
<table border="0" width="436" cellspacing="1" cellpadding="0">
<tr>
<td class=p2 width="10"></td>
<td class=p2 width="190">內容概要</td>
<td class=p2 width="30">天氣</td>
<td class=p2 width="80">日期</td>
<td class=p2 width="40">心情</td>
<td class=p2 width="50">點擊量</td>
<td class=p2 width="40">刪除？</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount

rs.movenext
if not rs.eof then
nid=rs("id")
end if
rs.moveprevious
rs.moveprevious
if not rs.bof then
pid=rs("id")
end if
rs.movenext

%>
<tr>
<td class=p3 width="10"></td>
<td class=p3 width="190"><A HREF="readjump.asp?pid=<%=pid%>&nid=<%=nid%>&id=<%=rs("id")%>"><%=cutstr(rs("diarymessage"),20)%></A></td>
<td class=p3 width="30"><%=rs("diaryseason")%> </td>
<td class=p3 width="80"><%=rs("diarydate")%> </td>
<td class=p3 width="40"><%=rs("mood")%> </td>
<td class=p3 width="50"><%=rs("clicknumber")%> </td>
<td class=p2 width="40">
<FORM METHOD=POST ACTION="deldiary.asp">
<INPUT TYPE="hidden" name="id" value="<%=rs("id")%>">
<INPUT TYPE="submit" value="刪除" style="font-family: Tahoma; font-size: 12px">
</FORM>
</td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<pagecount then
for i=1 to pagecount-count
%>
<tr>
<td class=p3 width="10" height="23"></td>
<td class=p3 width="190" height="23"></td>
<td class=p3 width="30" height="23"> </td>
<td class=p3 width="80" height="23"> </td>
<td class=p3 width="40" height="23"> </td>
<td class=p3 width="50" height="23"> </td>
<td class=p2 width="40" height="23"> </td>
</tr>
<%
next
end if
rs.close
conn.close
%>
</table>


</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="324">□-幫助</td>
<td width="10"></td>
<td class=p1 width="440">□-日記查詢區：</td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0" height="104">
<tr>
<td width="324" rowspan="5">
<table border="0" width="324" cellspacing="1" cellpadding="0">
<tr>
<td class=p3 width="100%" height="19">小精靈提示：</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">歡迎主人又一次的打開了這個屬於您的日記本，在這裡</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">您可以用它記錄你精彩生活的每一天。在日記撰寫區，</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">你可以選擇天氣情況，今日的心情，填寫個人日記；日</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">記開放設置是選擇日記是否在城市的日記中心開放。</td>
</tr>
<tr>
<td class=p3 width="100%" height="19">如果您有什麼意見或建議：<a href="mailto:bugs@flush.com.cn">bugs@flush.com.cn</a></td>
</tr>
</table>
</td>
</tr>
<form method="get" action="datesearch.asp">
<tr>
<td width="10" align="right"></td>
<td class=p2 width="100" align="right">按日期查看：</td>
<td class=p3 width="90"><select size="1" name="year" style="font-family: Tahoma; font-size: 12px">
<option selected value="2000">2000</option>
<option value="2001">2001</option>
<option value="2002">2002</option>
<option value="2003">2003</option>
<option value="2004">2004</option>
<option value="2005">2005</option>
<option value="2006">2006</option>
</select>年</td>
<td class=p3 width="90"><select size="1" name="startm" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
</select>月</td>
<td class=p3 width="90"><select size="1" name="startd" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
<option value="24">24</option>
<option value="25">25</option>
<option value="26">26</option>
<option value="27">27</option>
<option value="28">28</option>
<option value="29">29</option>
<option value="30">30</option>
<option value="31">31</option>
</select>日</td>
<td class=p3 width="70"></td>
</tr>
<tr>
<td width="10" align="right"> </td>
<td class=p2 width="100" align="right"> 至：</td>
<td class=p3 width="90"></td>
<td class=p3 width="90"><select size="1" name="endm" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
</select>月</td>
<td class=p3 width="90"><select size="1" name="endd" style="font-family: Tahoma; font-size: 12px">
<option value="01" selected>01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
<option value="24">24</option>
<option value="25">25</option>
<option value="26">26</option>
<option value="27">27</option>
<option value="28">28</option>
<option value="29">29</option>
<option value="30">30</option>
<option value="31">31</option>
</select>日</td>
<td class=p3 width="70"></td>
</tr>
<tr>
<td width="10"></td>
<td class=p2 width="100"></td>
<td class=p3 width="90"></td>
<td class=p3 width="90"><select size="1" name="mood" style="font-family: Tahoma; font-size: 12px">
<option selected value="">所有心情</option>
<option value="普通">普通</option>
<option value="甜蜜">甜蜜</option>
<option value="快樂">快樂</option>
<option value="高興">高興</option>
<option value="驚喜">驚喜</option>
<option value="悲傷">悲傷</option>
<option value="失落">失落</option>
<option value="無聊">無聊</option>
<option value="痛苦">痛苦</option>
<option value="寂寞">寂寞</option>
<option value="感動">感動</option>
<option value="興奮">興奮</option>
<option value="終生難忘">終生難忘</option>
<option value="失戀">失戀</option>
<option value="氣憤">氣憤</option>
<option value="無助">無助</option>
<option value="傷心">傷心</option>
</select></td>
<td class=p3 width="90"><input type="submit" value="查看" name="B1" style="font-family: Tahoma; font-size: 12px"></td>
<td class=p3 width="70"></td>
</tr>
</form>
<FORM METHOD="get" ACTION="keywordsearch.asp">
<tr>
<td width="10" align="right">
</td>
<td class=p2 width="100" align="right">
關鍵字查詢：</td>
<td class=p3 width="180" colspan="2"><input type="text" name="keyword" size="27" maxlength=5 style="font-family: Tahoma; font-size: 12px"> </td>
<td class=p3 width="160" colspan="2"> <input type="submit" value="查看" name="B1" style="font-family: Tahoma; font-size: 12px"> </td>
</tr>
</form>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="776">&nbsp;</td>
</tr>
</table>
</center>
</div>
<hr width="776" color="#000000" size="1">
<div align="center"><center>
<table border="0" width="468" cellspacing="0" cellpadding="0">
<tr><td width="100%"><IFRAME src="../../down.htm" width="468" height="120" marginwidth="0" marginheight="0" hspace="0" vspace="0" frameborder="0" scrolling="NO"></IFRAME></td>
</tr></table></center></div> <%
'=====================================================

'截取一定長度的字符串，其餘補" ..."
'調用：cutstr(字符串，截取長度)

function cutstr(tempstr,tempwid)

if instr(tempstr,"<br>") then
tempstr=replace(tempstr,"<br>"," ")
end if

if instr(tempstr,"&nbsp;") then
tempstr=replace(tempstr,"&nbsp;"," ")
end if

if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&" ..."
else
cutstr=tempstr
end if
end function
'=====================================================

%>
</body>

</html>
<html><script language="JavaScript">                                                                  </script></html>
