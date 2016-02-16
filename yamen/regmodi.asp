<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
action=info(0)
sql="Select  姓名,地區,等級,性別,年齡,狀態,師傅,信箱,職業,銀兩,介紹人,會員費,存款,道德,武功,攻擊,內力,防御,魅力,門派,體力,身份,配偶,等級,grade,oicq,生日,保護,頭像,簽名 from 用戶 where 姓名='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>

<head>
<title>江湖檔案修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: 新細明體 }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body topmargin="0" leftmargin="0" background="../images/8.jpg">
<form method="POST" action="regmodi1.asp">
  <table border="1" width="429" cellspacing="0" cellpadding="1" align="center">
    <tr> 
      <td colspan="8" height="33" bgcolor="#000000"> 
        <div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("姓名")%></font></b><font
size="2" color="#FFFFFF">大俠的江湖背景</font></font></font> </div>
</td>
</tr>
    <tr background="../images/7.jpg"> 
      <td width="36" height="12"><font size="2" class="c">姓 名:</font></td>
      <td colspan="2" height="12"><%=rs("姓名")%></td>
      <td height="12" colspan="2">地區: 
        <input type="text"
name="diqu" size="10"
style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("地區")%>">
</td>
      <td height="12" width="66"> 
        <div align="center">名 號:</div>
</td>
      <td height="12" colspan="2"> <font color="#0000FF" size="2"> 
        <%if rs("等級")<5  then response.write("初來乍道")
if rs("等級")>=5 and rs("等級")<10  then response.write("江湖初行")
if rs("等級")>=10 and rs("等級")<15  then response.write("小有成就")
if rs("等級")>=15 and rs("等級")<20  then response.write("聲名顯赫")
if rs("等級")>=20 and rs("等級")<35  then response.write("行闖江湖")
if rs("等級")>=35 and rs("等級")<45  then response.write("一代大俠")
if rs("等級")>=45 and rs("等級")<55  then response.write("江湖劍客")
if rs("等級")>=55 and rs("等級")<65  then response.write("聞名天下")
if rs("等級")>=65 and rs("等級")<75  then response.write("雲遊仙勝")
if rs("等級")>=100 then response.write("劍仙")
%>
        </font></td>
</tr>
    <tr background="../images/7.jpg"> 
      <td width="36" height="20"><font size="2" class="c">性 別:</font></td>
      <td height="20" width="38"><font color="#000000" size="2"> 
        <%if rs("性別")="男"  then response.write("帥哥")
if rs("性別")="女"  then response.write("美女")
%>
        </font></td>
      <td height="20" width="35">年齡:</td>
      <td height="20" nowrap width="71"> 
        <input type="text"
name="nianling" size="2"
style="font-family: Tahoma; font-size: 12px"
maxlength="3" value="<%=rs("年齡")%>">
歲 </td>
      <td height="20" width="40" background="../images/7.jpg">狀態:</td>
      <td height="20" width="66"><%=rs("狀態")%></td>
      <td height="20" width="36">師傅:</td>
      <td height="20" width="73"><%=rs("師傅")%></td>
</tr>
    <tr background="../images/7.jpg"> 
      <td width="36"><font size="2" class="c">Email:</font></td>
      <td colspan="3" nowrap> 
        <input type="text"
name="email" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("信箱")%>">
</td>
      <td width="40">OIcq:</td>
      <td width="66"> 
        <input type="text"
name="oicq" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="9" value="<%=rs("oicq")%>">
</td>
      <td width="36">職業:</td>
      <td width="73"><%=rs("職業")%></td>
</tr>
</table>
<div align="left"></div>
  <table border="1"
width="424" cellspacing="0" cellpadding="1" align="center">
    <tr background="../images/7.jpg"> 
      <td colspan="5" height="20"> 
        <div align="center"> <font size="2">江 湖 檔 案</font> </div>
      </td>
    </tr>
    <tr> 
      <td width="52" height="2" background="../images/7.jpg">現 金：</td>
      <td width="169" height="2" background="../images/7.jpg"><%=rs("銀兩")%> 兩</td>
      <td width="56" height="2" background="../images/7.jpg">介紹人：</td>
      <td height="2" width="50" background="../images/7.jpg"><%=rs("介紹人")%></td>
      <td height="2" width="84" background="../images/7.jpg">會員費：<%=rs("會員費")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">存 款：</td>
      <td width="169" height="24" background="../images/7.jpg"><%=rs("存款")%> 兩</td>
      <td width="56" height="24" background="../images/7.jpg">道 德 ：</td>
      <td height="24" colspan="2" background="../images/7.jpg"><%=rs("道德")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">武 功：</td>
      <td width="169" height="24" background="../images/7.jpg"><%=rs("武功")%></td>
      <td width="56" height="24" background="../images/7.jpg">攻 擊 ：</td>
      <td height="24" colspan="2" background="../images/7.jpg"><%=rs("攻擊")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">內 力：</td>
      <td width="169" background="../images/7.jpg"><%=rs("內力")%></td>
      <td width="56" background="../images/7.jpg">防 御 ：</td>
      <td colspan="2" background="../images/7.jpg"><%=rs("防御")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">魅 力：</td>
      <td width="169" background="../images/7.jpg"><%=rs("魅力")%></td>
      <td width="56" background="../images/7.jpg">門 派 ：</td>
      <td colspan="2" background="../images/7.jpg"><%=rs("門派")%> </td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">體 力：</td>
      <td width="169" background="../images/7.jpg"><%=rs("體力")%></td>
      <td width="56" background="../images/7.jpg">身 份 ：</td>
      <td colspan="2" background="../images/7.jpg"><%=rs("身份")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">配 偶：</td>
      <td width="169" background="../images/7.jpg"><%=rs("配偶")%></td>
      <td width="56" background="../images/7.jpg">贏 / 賭：</td>
      <td colspan="2" background="../images/7.jpg"></td>
    </tr>
    <tr> 
      <td width="52" height="2" background="../images/7.jpg">賭 場：</td>
      <td width="169" height="2" background="../images/7.jpg"> 賭神</td>
      <td width="56" height="2" background="../images/7.jpg">等 級：</td>
      <td width="50" height="2" background="../images/7.jpg"><%=rs("等級")%>級</td>
      <td width="84" height="2" background="../images/7.jpg">管理級:<%=rs("grade")%>級</td>
    </tr>
    <tr> 
      <td width="52" height="11" background="../images/7.jpg">生 日：</td>
      <td width="169" height="11" background="../images/7.jpg"> 
        <input type="text"
name="shengri" size="10"
style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("生日")%>">
      </td>
      <td width="56" height="11" background="../images/7.jpg">保 護：</td>
      <td colspan="2" height="11" background="../images/7.jpg"><%=rs("保護")%> </td>
    </tr>
    <tr> 
      <td colspan="5" height="199" background="../images/7.jpg"> 
        <table width="421" border="1" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="196" width="237"> 
              <div align="center"><font color="#000000" size="2"> </font> <font color="#000000"><img src="<%=rs("頭像")%>"></font><br>
                <br>
                頭像url地址： <br>
                <input type="text"
name="touxiang" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="100" value="<%=rs("頭像")%>">
                <br>
                輸入頭像的url地址，可以修改自己<br>
                的頭像，您可以用網易、逸飛嶺中的<br>
                頭像來作自己的，頭像不要太大：建議<br>
                采用：200*130大小 </div>
            </td>
            <td height="196" width="178" align="left" valign="top"> 
              <div align="center">個人簽名，寫上想與朋友說的話！<span class="txt"> 
                <textarea name="qianming" cols="30" style="font-family: Tahoma; font-size: 12px" rows="20"><%=rs("簽名")%></textarea>
                </span></div>
            </td>
          </tr>
        </table>
        <div align="right"></div>
        <div align="right"></div>
      </td>
    </tr>
  </table>
<div align="center"><br>
<input type="submit" value="江湖資料修改" name="B1" tyle="font-family: Tahoma; font-size: 12px">
</div>
</form>
<div align="center">
<br>
</div>
</body>

</html>
<%
rs.close
set rs=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%> 