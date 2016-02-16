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
<%Response.Expires=0
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from ゅセ "
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Redirect "modifygg.asp"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#0033FF" link="#0000FF" vlink="#0000FF" alink="#0000FF" background="0064.jpg">
<div align="center"><font color="#FF0000" size="+6">册ぱ簎笆約穦よΑэ</font></div>
<form method="post" action="modifyggok.asp">
  <table border="1" cellspacing="0" align="center" cellpadding="0" width="560" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
    <tr bgcolor="#99CCFF"> 
      <td colspan="2"> 
        <p><font size="2" color="#FF0000">猔種计沮畐穝祘丁Τ⊿Τ浪代叫產э猔種⊿Τよ恶0<br>
          惠璶ゴó块才&lt;br&gt;<br>
          穦る1234癸莱穦虏ざい1234穦摸崩Τぃ矪叫筿秎:Seven.s-888@yahoo.com.tw</font> </p>
        </td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">id</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="gg" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" readonly
value="<%=rs("id")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">簎笆約</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="gg1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("簎笆約")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">1穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy1"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る1")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">2穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy2"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る2")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">3穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy3"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る3")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">4穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy4"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る4")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">5穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy17"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る5")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">6穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy18"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る6")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">7穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy19"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る7")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">8穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy20"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る8")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">9穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy21"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る9")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">10穦る</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy22"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦る10")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">1穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy5"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单1")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">2穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy6"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单2")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">3穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy7"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单3")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">4穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy8"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单4")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">5穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy23"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单5")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">6穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy24"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单6")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">7穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy25"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单7")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">8穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy26"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单8")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">9穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy27"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单9")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">10穦单</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy28"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦单10")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">獶穦獁翴计</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy91"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("獶穦獁翴")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">1穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy9"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴1")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">2穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy10"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴2")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">3穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy11"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴3")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">4穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy12"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴4")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">5穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy29"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴5")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">6穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy30"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴6")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">7穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy31"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴7")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">8穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy32"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴8")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">9穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy33"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴9")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">10穦獁翴</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy34"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦獁翴10")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">穦蹿纒</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy13"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦纒")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">穦蹿</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy14"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("穦")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600"></font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy15"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">PK痁</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy16"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("PK")%>" size="60">
      </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
        <input type="submit" value="絋 ﹚">
        <font color="#CCCCCC">------- </font> 
        <input  onClick="javascript:window.document.location.href='modifygg.asp'" value=" 穝" type=button name="back">
        </div>
    </td>
    </tr>
  </table>
</form>
<%
end if
set rs=nothing	
	conn.close
	set conn=nothing
%>
<div align="center"> </div>
