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
action=request.querystring("action")
if action="大家" then action=info(0)
gw=left(action,1)
if IsNumeric(gw)=true then
zz=split(action,"級")
gw=zz(0)
Response.Write "<script language=JavaScript>{alert('怪物信息不可見！');window.close();}</script>"
	Response.End
end if
if info(0)<>action then
sql="Select  銀兩,grade from 用戶 where 姓名='"& info(0) &"'"
set rs=conn.Execute(sql)
if rs("銀兩")<10000 and rs("grade")<7 then%>
<script language=vbscript>
MsgBox "你沒有一萬兩銀子打點查不到東西的！"
window.close()
</script>
<%response.end
end if
sql="update 用戶 set 銀兩=銀兩-10000 where 姓名='"& info(0) &"'"
Set Rs=conn.Execute(sql)
end if
sql="Select  id,姓名,地區,等級,性別,年齡,狀態,師傅,職業,銀兩,介紹人,會員費,存款,道德,武功,攻擊,金幣,內力,防御,魅力,寶物,門派,體力,法力,身份,配偶,二婚,等級,grade,頭像,簽名 from 用戶 where 姓名='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>
<head>
<title>江湖檔案</title>
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

<body topmargin="0" leftmargin="0" bgcolor="#006699" text="#FFFFFF" link="#FFFFCC" vlink="#FFFFCC" alink="#FFFFCC">
<table border="1"
width="429" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
  <tr bgcolor="#333366"> 
    <td colspan="9" height="33"> 
      <div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF">ID:</font><font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("id")%> 
        </font></b></font></font><font color="#FFFFFF"><%=rs("姓名")%></font></b><font
size="2" color="#FFFFFF">大俠的江湖背景</font></font></font> </div>
    </td>
  </tr>
  <tr bgcolor="#003366"> 
    <td width="40" height="12"><font size="2" class="c">姓 名:</font></td>
    <td colspan="2" height="12"><%=rs("姓名")%></td>
    <td height="12" colspan="3">地區:<%=rs("地區")%></td>
    <td height="12" width="48"> 
      <div align="center">名 號:</div>
    </td>
    <td height="12" colspan="2"> <font color="#FFFFCC" size="2"> 
      <%
if info(0)=action then
zddj=-1
else
zddj=rs("等級")
end if
if rs("等級")<5  then response.write("初來乍道")
if rs("等級")>=5 and rs("等級")<10  then response.write("江湖初行")
if rs("等級")>=10 and rs("等級")<15  then response.write("小有成就")
if rs("等級")>=15 and rs("等級")<20  then response.write("聲名顯赫")
if rs("等級")>=20 and rs("等級")<35  then response.write("行闖江湖")
if rs("等級")>=35 and rs("等級")<45  then response.write("一代大俠")
if rs("等級")>=45 and rs("等級")<55  then response.write("江湖劍客")
if rs("等級")>=55 and rs("等級")<65  then response.write("聞名天下")
if rs("等級")>=65 and rs("等級")<75  then response.write("雲遊仙勝")
if rs("等級")>=75 and rs("等級")<80  then response.write("已入仙道")
if rs("等級")>=80 and rs("等級")<85  then response.write("小仙")
if rs("等級")>=85 and rs("等級")<90  then response.write("大仙")
if rs("等級")>=90 and rs("等級")<95  then response.write("極帝大仙")
if rs("等級")>=95 and rs("等級")<100  then response.write("仙人")
if rs("等級")>=100 then response.write("劍仙")
%>
      </font></td>
  </tr>
  <tr bgcolor="#003366"> 
    <td width="40" height="14"><font size="2" class="c">性 別:</font></td>
    <td height="14" colspan="2"><font color="#FFFFCC" size="2"> 
      <%if rs("性別")="男"  then response.write("帥哥")
if rs("性別")="女"  then response.write("美女")
%>
      </font></td>
    <td height="14" nowrap width="47">年齡:</td>
    <td height="14" nowrap width="48"><%=rs("年齡")%> 歲</td>
    <td height="14" width="40">狀態:</td>
    <td height="14" width="48"><%=rs("狀態")%></td>
    <td height="14" width="49">師傅:</td>
    <td height="14" width="60"><%=rs("師傅")%></td>
  </tr>
  <tr bgcolor="#003366"> 
    <td width="40">職業:</td>
    <td width="32"><%=rs("職業")%></td>
    <td colspan="5"><font color="#FFFFFF"></font></td>
    <td width="49">金幣：</td>
    <td width="60"><font color="#FFFFFF"><%=rs("金幣")%><font color="#FFFFCC">個</font></font></td>
  </tr>
</table>
<div align="left"></div>
<table border="1"
width="429" cellspacing="0" cellpadding="1" align="center" bgcolor="#003366">
  <tr> 
    <td colspan="6" height="20"> 
      <div align="center"> <font size="2">江 湖 檔 案</font> </div>
    </td>
  </tr>
  <tr> 
    <td width="60" height="25">現 金：</td>
    <td height="25" colspan="2"> <%=rs("銀兩")%> 兩 </td>
    <td width="62" height="25">介紹人：</td>
    <td height="25" width="52"> <%=rs("介紹人")%> </td>
    <td height="25" width="86">會費： <font color="#FFFFFF"><%=rs("會員費")%><font color="#FFFFCC">元</font></font> 
    </td>
  </tr>
  <tr> 
    <td width="60" height="20">存 款：</td>
    <td height="24" colspan="2"> <%=rs("存款")%> 兩 </td>
    <td width="62" height="24">道 德：</td>
    <td height="24" colspan="2"> <%=rs("道德")%> </td>
  </tr>
  <tr> 
    <td width="60" height="20">武 功：</td>
    <td height="24" colspan="2"> <%=rs("武功")%> </td>
    <td width="62" height="24">攻 擊：</td>
    <td height="24" colspan="2"> <%=rs("攻擊")%> </td>
  </tr>
  <tr> 
    <td width="60" height="20">內 力：</td>
    <td colspan="2"> <%=rs("內力")%> </td>
    <td width="62">防 御：</td>
    <td colspan="2"> <%=rs("防御")%> </td>
  </tr>
  <tr> 
    <td width="60" height="20">魅 力：</td>
    <td colspan="2"> <%=rs("魅力")%> </td>
    <td width="62">門 派：</td>
    <td width="52"><%=rs("門派")%> </td>
    <td width="86">寶物:<%=rs("寶物")%></td>
  </tr>
  <tr> 
    <td width="60" height="20">體 力：</td>
    <td width="72"> <%=rs("體力")%> </td>
    <td width="72">法力：<%=rs("法力")%></td>
    <td width="62">身 份：</td>
    <td> <%=rs("身份")%> </td>
    
  </tr>
  <tr> 
    <td width="60" height="20">配 偶：</td>
    <td colspan="2"> <%=rs("配偶")%> 二婚：<%=rs("二婚")%></td>
    <td width="62">贏 / 賭：</td>
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td width="60" height="20">賭場等級：</td>
    <td colspan="2"> 
    </td>
    <td width="62">江湖等級：</td>
    <td width="52"><%=rs("等級")%>級</td>
    <td width="86">管理等級:<%=rs("grade")%>級</td>
  </tr>
  <tr> 
    <td colspan="6" height="199"> 
      <table width="424" border="1" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="196" width="223"> 
            <div align="center"><font color="#000000" size="2"> </font><font color="#000000"><img src="<%=rs("頭像")%>"></font></div>
          </td>
          <td height="196" width="184" align="left" valign="top">個人簽名：<span class="txt"><%=changechr(rs("簽名"))%></span></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>

</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%> 