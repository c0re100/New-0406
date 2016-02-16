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

<!--#include file="data2.asp"-->
<html>

<head>
<title>靈劍江湖-注冊詳細資料</title>
<link rel="stylesheet" type="text/css" href="../../style.css">
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script language="JavaScript">
<!--
function view(newsfile){var gt=unescape('%3e');var popup=null;var over="Launch Pop-up Navigator";popup=window.open('','popupnav','width=510,height=470,resizable=1,scrollbars=yes');if(popup!=null){if(popup.opener==null){popup.opener=self;}popup.location.href=newsfile;}}//-->
<!--
function GotoURL(value)
{
url=value
location.href=url
}
// -->
</script>

</head>

<body bgcolor="#0099FF" text="#FFFF00" background="../../jhimg/BK_HC3W.GIF">
<%
rs.open"select * from userinfo where user='"&info(0)&"'",conn,1,1

skin=rs("skin")
if skin<>"" then
skin=rs("skin")
else
skin="01.gif"
end if

rs.close
rs.open"select * from userinfo where user='"&info(0)&"'",conn,1,1
%>
<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
<tr>
    <td width="744" height="83"> 
      <table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
align="center">
<tbody>
<tr>
<td height="38" width="100%">
<table width="100%" border="0" cellspacing="0" cellpadding="0"
bordercolorlight="#000000" bordercolordark="#FFFFFF"
align="center">
<tr>
<td width="91" height="26"><font size="2">&nbsp; <font
color="#FFFFFF"></font><font size="2"><font color="#ffffff"
size="2"><span class="zilong"><font color="#FFCC00">
<%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %>
</font></span></font></font></font></td>
<td width="475" height="26">
<div align="center">
<font size="2" color="#E18C59"><b>填寫詳細資料</b></font>
</div>
</td>
<td width="104">
<div align="center">
<a href="../../jh.asp" target="_top"><font color="#E18C59">返回皇城中心</font></a>
</div>


</td>
</tr>
</table>
</td>
</tr>
</tbody>
</table>
    </td>
</tr>
</table>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
    <td width="17">&nbsp;</td>
<td valign="top" width="639">
<div align="center">
<div align="center">
<div align="center">
<table border="0" width="468" cellspacing="1" cellpadding="0"
height="20">

</table>
</div>
</div>
<div align="center">
<center>
<table border="0" width="695" cellspacing="1" cellpadding="0">
<tr>
<td width="690">
<form method="POST" action="reg.asp">
<div align="center">
<center>
<table border="1" width="690" cellspacing="1"
cellpadding="0" bordercolor="#E18C59">
<tr>
<td width="496" class="p2" colspan="4">如果您想成為皇城的正式居民（有身份證的：），請您認真填寫以下信息：</td>
<td width="13" rowspan="23"> </td>
</tr>
<tr>
<td width="69" class="p2">真實姓名：</td>
<td width="196" class="p2"><input type="text"
name="name" size="28"
style="font-family: Tahoma; font-size: 12px"
value="<%=rs("name")%>" maxlength="20"> <font
color="#FF0000">*</font></td>
<td width="71" class="p2">真實性別：</td>
<td width="156" class="p2"><select name="sex" size="1"
style="font-family: Tahoma; font-size: 12px">
<option value="男"
<%
if instr(rs("sex"),"男") then
response.write "selected"
end if
%>>男</option>
<option value="女"
<%
if instr(rs("sex"),"女") then
response.write "selected"
end if
%>>女</option>
</select></td>
</tr>
<tr>
<td width="69" class="p3">出生日期：</td>
<td width="423" class="p3" colspan="3"><select size="1"
name="birthyear"
style="font-family: Tahoma; font-size: 12px">
<%for i = 1900 to 2000%>
<option value="<%=i%>"
<%
if int(left(rs("birthday"),4))=i then
response.write "selected"
end if
%>><%=i%></option>
<%next%>
</select> 年 <select size="1" name="birthmonth"
style="font-family: Tahoma; font-size: 12px">
<option value="01" <%
if mid(rs("birthday"),6,2)="01" then
response.write "selected"
end if
%>>01</option>
<option value="02"
<%
if mid(rs("birthday"),6,2)="02" then
response.write "selected"
end if
%>>02</option>
<option value="03"
<%
if mid(rs("birthday"),6,2)="03" then
response.write "selected"
end if
%>>03</option>
<option value="04"
<%
if mid(rs("birthday"),6,2)="04" then
response.write "selected"
end if
%>>04</option>
<option value="05"
<%
if mid(rs("birthday"),6,2)="06" then
response.write "selected"
end if
%>>05</option>
<option value="06"
<%
if mid(rs("birthday"),6,2)="06" then
response.write "selected"
end if
%>>06</option>
<option value="07"
<%
if mid(rs("birthday"),6,2)="07" then
response.write "selected"
end if
%>>07</option>
<option value="08"
<%
if mid(rs("birthday"),6,2)="08" then
response.write "selected"
end if
%>>08</option>
<option value="09"
<%
if mid(rs("birthday"),6,2)="09" then
response.write "selected"
end if
%>>09</option>
<option value="10"
<%
if mid(rs("birthday"),6,2)="10" then
response.write "selected"
end if
%>>10</option>
<option value="11"
<%
if mid(rs("birthday"),6,2)="11" then
response.write "selected"
end if
%>>11</option>
<option value="12"
<%
if mid(rs("birthday"),6,2)="12" then
response.write "selected"
end if
%>>12</option>
</select> 月 <select size="1" name="birthdate"
style="font-family: Tahoma; font-size: 12px">
<%for j = 1 to 31%>
<option
value="<%
if j<10 then
putday="0"&trim(cstr(j))
else
putday=trim(cstr(j))
end if
response.write putday
%>"
<%
if int(mid(rs("birthday"),9,2))=j then
response.write "selected"
end if
%>><%=putday%></option>
<%next%>
</select> 日 <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">電子郵箱：</td>
<td width="200" class="p2"><input type="text"
name="email" size="35"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("email")%>"> <font
color="#FF0000">*</font></td>
<td width="73" class="p2" rowspan="2" align="center"><img
src="skin/<%=skin%>" alt="<%=session("user")%>"
width="40" height="40"></td>
<td width="158" class="p2"><a href="photo.asp"
target="_blank"><font color="#FFCC00">更改我的頭像</font></a></td>
</tr>
<tr>
<td width="69" class="p3">個人主頁：</td>
<td width="200" class="p3"><input type="text"
name="homepage" size="35"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("homepage")%>"> <font
color="#FF0000">*</font></td>
<td width="158" class="p3"> </td>
</tr>
<tr>
<td width="69" class="p2">家庭地址：</td>
<td width="423" class="p2" colspan="3"><input
type="text" name="address" size="78"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("address")%>"></td>
</tr>
<tr>
<td width="69" class="p3">郵政編碼：</td>
<td width="194" class="p3"><input type="text" name="pc"
size="23" style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("pc")%>"> <font
color="#FF0000">*</font></td>
<td width="68" class="p3">聯繫電話：</td>
<td width="153" class="p3"><input type="text" name="tel"
size="23" style="font-family: Tahoma; font-size: 12px"
maxlength="25" value="<%=rs("tel")%>"> <font
color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">職業：</td>
<td width="194" class="p2"><input type="text" name="job"
size="23" style="font-family: Tahoma; font-size: 12px"
maxlength="25" value="<%=rs("job")%>"></td>
<td width="68" class="p2">身份證號：</td>
<td width="153" class="p2"><input type="text"
name="idcard" size="23"
style="font-family: Tahoma; font-size: 12px"
maxlength="20" value="<%=rs("idcard")%>"> <font
color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p3">工作單位：</td>
<td width="423" class="p3" colspan="3"><input
type="text" name="corporation" size="78"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("corporation")%>"></td>
</tr>
<tr>
<td width="69" class="p2">月收入：</td>
<td width="194" class="p2"><select size="1"
name="salary"
style="font-family: Tahoma; font-size: 12px">
<option value="無收入" selected>無收入</option>
<option value="500元以下">500元以下</option>
<option value="500–1000元">500–1000元</option>
<option value="1000–2000元">1000–2000元</option>
<option value="2000–4000元">2000–4000元</option>
<option value="4000–8000元">4000–8000元</option>
<option value="8000元以上">8000元以上</option>
</select></td>
<td width="68" class="p2"> </td>
<td width="153" class="p2"> </td>
</tr>
<tr>
<td width="69" class="p3">上網方式：</td>
<td width="423" class="p3" colspan="3">
<table border="0" width="100%" cellspacing="1"
cellpadding="0">
<tr>
<td width="6%"><input type="checkbox" name="mode1"
value="家中"
<%
if instr(rs("mode"),"家中") then
response.write "checked"
end if
%>></td>
<td width="26%">家中</td>
<td width="6%"><input type="checkbox" name="mode2"
value="單位"
<%
if instr(rs("mode"),"單位") then
response.write "checked"
end if
%>></td>
<td width="28%">單位</td>
<td width="6%"><input type="checkbox" name="mode3"
value="網吧"
<%
if instr(rs("mode"),"網吧") then
response.write "checked"
end if
%>></td>
<td width="28%">網吧</td>
</tr>
</table>
</td>
</tr>
<tr>
<td width="69" class="p2">身高：</td>
<td width="194" class="p2"><input type="text"
name="stature" size="11"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("stature")%>"> 釐米 <font
color="#FF0000">*</font></td>
<td width="68" class="p2">婚姻狀況：</td>
<td width="153" class="p2"><select size="1"
name="marriage"
style="font-family: Tahoma; font-size: 12px">
<option value="未婚"
<%
if instr(rs("marriage"),"未婚") then
response.write "selected"
end if
%>>未婚</option>
<option value="已婚"
<%
if instr(rs("marriage"),"已婚") then
response.write "selected"
end if
%>>已婚</option>
<option value="離婚"
<%
if instr(rs("marriage"),"離婚") then
response.write "selected"
end if
%>>離婚</option>
<option value="喪偶"
<%
if instr(rs("marriage"),"喪偶") then
response.write "selected"
end if
%>>喪偶</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p3">體重：</td>
<td width="194" class="p3"><input type="text"
name="weight" size="7"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("weight")%>"> 公斤 <font
color="#FF0000">*</font></td>
<td width="68" class="p3">所在城市：</td>
<td width="153" class="p3"><select size="1" name="city"
style="font-family: Tahoma; font-size: 12px">
<option value="四川"
<%
if instr(rs("city"),"四川") then
response.write "selected"
end if
%>>四川</option>
<option value="天津"
<%
if instr(rs("city"),"天津") then
response.write "selected"
end if
%>>天津</option>
<option value="北京"
<%
if instr(rs("city"),"北京") then
response.write "selected"
end if
%>>北京</option>
<option value="廣東"
<%
if instr(rs("city"),"廣東") then
response.write "selected"
end if
%>>廣東</option>
<option value="上海"
<%
if instr(rs("city"),"上海") then
response.write "selected"
end if
%>>上海</option>
<option value="廣西"
<%
if instr(rs("city"),"廣西") then
response.write "selected"
end if
%>>廣西</option>
<option value="海南"
<%
if instr(rs("city"),"海南") then
response.write "selected"
end if
%>>海南</option>
<option value="湖南"
<%
if instr(rs("city"),"湖南") then
response.write "selected"
end if
%>>湖南</option>
<option value="湖北"
<%
if instr(rs("city"),"湖北") then
response.write "selected"
end if
%>>湖北</option>
<option value="江西"
<%
if instr(rs("city"),"江西") then
response.write "selected"
end if
%>>江西</option>
<option value="江蘇"
<%
if instr(rs("city"),"江蘇") then
response.write "selected"
end if
%>>江蘇</option>
<option value="浙江"
<%
if instr(rs("city"),"浙江") then
response.write "selected"
end if
%>>浙江</option>
<option value="福建"
<%
if instr(rs("city"),"福建") then
response.write "selected"
end if
%>>福建</option>
<option value="安徽"
<%
if instr(rs("city"),"安徽") then
response.write "selected"
end if
%>>安徽</option>
<option value="雲南"
<%
if instr(rs("city"),"雲南") then
response.write "selected"
end if
%>>雲南</option>
<option value="貴州"
<%
if instr(rs("city"),"貴州") then
response.write "selected"
end if
%>>貴州</option>
<option value="重慶"
<%
if instr(rs("city"),"重慶") then
response.write "selected"
end if
%>>重慶</option>
<option value="河南"
<%
if instr(rs("city"),"河南") then
response.write "selected"
end if
%>>河南</option>
<option value="河北"
<%
if instr(rs("city"),"河北") then
response.write "selected"
end if
%>>河北</option>
<option value="山東"
<%
if instr(rs("city"),"山東") then
response.write "selected"
end if
%>>山東</option>
<option value="山西"
<%
if instr(rs("city"),"山西") then
response.write "selected"
end if
%>>山西</option>
<option value="遼寧"
<%
if instr(rs("city"),"遼寧") then
response.write "selected"
end if
%>>遼寧</option>
<option value="吉林"
<%
if instr(rs("city"),"吉林") then
response.write "selected"
end if
%>>吉林</option>
<option value="黑龍江"
<%
if instr(rs("city"),"黑龍江") then
response.write "selected"
end if
%>>黑龍江</option>
<option value="陝西"
<%
if instr(rs("city"),"陝西") then
response.write "selected"
end if
%>>陝西</option>
<option value="青海"
<%
if instr(rs("city"),"青海") then
response.write "selected"
end if
%>>青海</option>
<option value="西藏"
<%
if instr(rs("city"),"西藏") then
response.write "selected"
end if
%>>西藏</option>
<option value="新疆"
<%
if instr(rs("city"),"新疆") then
response.write "selected"
end if
%>>新疆</option>
<option value="甘肅"
<%
if instr(rs("city"),"甘肅") then
response.write "selected"
end if
%>>甘肅</option>
<option value="寧夏"
<%
if instr(rs("city"),"寧夏") then
response.write "selected"
end if
%>>寧夏</option>
<option value="內蒙古"
<%
if instr(rs("city"),"內蒙古") then
response.write "selected"
end if
%>>內蒙古</option>
<option value="臺灣"
<%
if instr(rs("city"),"臺灣") then
response.write "selected"
end if
%>>臺灣</option>
<option value="9港"
<%
if instr(rs("city"),"9港") then
response.write "selected"
end if
%>>9港</option>
<option value="澳門"
<%
if instr(rs("city"),"澳門") then
response.write "selected"
end if
%>>澳門</option>
<option value="其它"
<%
if instr(rs("city"),"其它") then
response.write "selected"
end if
%>>其它</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">民族：</td>
<td width="194" class="p2"><input type="text"
name="nation" size="7"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("nation")%>"> 族 <font
color="#FF0000">*</font></td>
<td width="68" class="p2">學歷：</td>
<td width="153" class="p2"><select name="education"
size="1" style="font-family: Tahoma; font-size: 12px">
<option value="大學"
<%
if instr(rs("education"),"大學") then
response.write "selected"
end if
%>>大學</option>
<option value="小學"
<%
if instr(rs("education"),"小學") then
response.write "selected"
end if
%>>小學</option>
<option value="初中"
<%
if instr(rs("education"),"初中") then
response.write "selected"
end if
%>>初中</option>
<option value="高中"
<%
if instr(rs("education"),"高中") then
response.write "selected"
end if
%>>高中</option>
<option value="碩士"
<%
if instr(rs("education"),"碩士") then
response.write "selected"
end if
%>>碩士</option>
<option value="博士"
<%
if instr(rs("education"),"博士") then
response.write "selected"
end if
%>>博士</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p3">ICQ號碼：</td>
<td width="194" class="p3"><input type="text" name="icq"
size="19" style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("icq")%>"></td>
<td width="68" class="p3">健康狀況：</td>
<td width="153" class="p3"><select size="1"
name="health"
style="font-family: Tahoma; font-size: 12px">
<option value="健康"
<%
if instr(rs("health"),"健康") then
response.write "selected"
end if
%>>健康</option>
<option value="良好"
<%
if instr(rs("health"),"良好") then
response.write "selected"
end if
%>>良好</option>
<option value="一般"
<%
if instr(rs("health"),"一般") then
response.write "selected"
end if
%>>一般</option>
<option value="差"
<%
if instr(rs("health"),"差") then
response.write "selected"
end if
%>>差</option>
</select> <font color="#FF0000">*</font></td>
</tr>
<tr>
<td width="69" class="p2">OICQ號碼：</td>
<td width="194" class="p2"><input type="text"
name="oicq" size="19"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("oicq")%>"></td>
<td width="68" class="p2"> </td>
<td width="153" class="p2"> </td>
</tr>
<tr>
<td width="69" class="p3">尋呼：</td>
<td width="194" class="p3"><input type="text"
name="pager" size="19"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("pager")%>"></td>
<td width="68" class="p3"> </td>
<td width="153" class="p3"> </td>
</tr>
<tr>
<td width="69" class="p2">移動電話：</td>
<td width="194" class="p2"><input type="text"
name="mobilephone" size="19"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("mobilephone")%>"></td>
<td width="68" class="p2"> </td>
<td width="153" class="p2"> </td>
</tr>
<tr>
<td width="69" class="p2" valign="top">個人愛好：</td>
<td width="415" class="p2" colspan="3"><textarea
rows="3" name="preference" cols="81"
style="font-family: Tahoma; font-size: 12px"><%=rs("preference")%></textarea>
</tr>
<tr>
<td width="69" class="p2" valign="top">個人介紹：<br>
<font color="#FF0000">*</font></td>
<td width="415" class="p2" colspan="3"><textarea
rows="7" name="introduction" cols="81"
style="font-family: Tahoma; font-size: 12px"><%=rs("introduction")%></textarea></td>
</tr>
<tr>
<td width="69" class="p3" valign="top">是否開放您的個人信息</td>
<td width="415" class="p3" colspan="3"><input
type="radio" value="是" name="open"
<%
if instr(rs("open"),"是") then
response.write "checked"
end if
%>> 是 <input type="radio" value="否" name="open"
<%
if instr(rs("open"),"否") then
response.write "checked"
end if
%>> 否</td>
</tr>
<tr>
<td width="69" class="p2" valign="top"> </td>
<td width="415" class="p2" colspan="3"><input
type="submit" value="修改" name="B1"
style="font-family: Tahoma; font-size: 12px"></td>
</tr>
</table>
</center>
</div>
</form>
</td>
</tr>
</table>
</center>
</div>
</div>
</td>
    <td width="25"> </td>
</tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td>
      <div align="right"> </div>
</td>
</tr>
</table>
<div align="center">
</div>

</body>

</html>
<%
rs.close
conn.close
%>