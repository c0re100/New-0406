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
sql="select * from �奻 "
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
<div align="center"><font color="#FF0000" size="+6">��ѫǺu�ʼs�i�B�|���覡�ק�</font></div>
<form method="post" action="modifyggok.asp">
  <table border="1" cellspacing="0" align="center" cellpadding="0" width="560" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
    <tr bgcolor="#99CCFF"> 
      <td colspan="2"> 
        <p><font size="2" color="#FF0000">�`�N�G�ƾڮw��s�{�Ǧ]���ɶ������A�S���[�J���ŭȮɪ��˴��A�Фj�a���ɪ`�N�S���Ȫ��a���0<br>
          �ݭn���^������J�r��&lt;br&gt;<br>
          �|����I1�B2�B3�B4������|��²������1�šB2�šB3�šB4�ŷ|���A�H�������A�p�������B�йq�l:Seven.s-888@yahoo.com.tw</font> </p>
        </td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">id</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="gg" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" readonly
value="<%=rs("id")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">�u�ʼs�i</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="gg1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�u�ʼs�i")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">1�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy1"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I1")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">2�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy2"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I2")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">3�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy3"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I3")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">4�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy4"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I4")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">5�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy17"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I5")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">6�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy18"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I6")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">7�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy19"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I7")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">8�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy20"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I8")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">9�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy21"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I9")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">10�ŷ|����I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy22"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|����I10")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">1�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy5"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������1")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">2�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy6"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������2")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">3�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy7"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������3")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">4�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy8"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������4")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">5�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy23"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������5")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">6�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy24"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������6")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">7�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy25"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������7")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">8�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy26"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������8")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">9�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy27"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������9")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">10�ŷ|������</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy28"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|������10")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">�D�|���w�I��</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy91"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�D�|���w�I")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">1�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy9"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I1")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">2�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy10"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I2")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">3�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy11"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I3")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">4�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy12"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I4")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">5�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy29"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I5")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">6�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy30"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I6")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">7�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy31"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I7")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">8�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy32"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I8")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">9�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy33"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I9")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">10�ŷ|���w�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy34"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���w�I10")%>" size="20">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">�|���I���s�d</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy13"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���s�d")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">�|���I�ڦa�}</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy14"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�|���a�}")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">�b�~�I</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy15"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("�b�~�I")%>" size="60">
      </font></td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">PK�ȯZ��</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="hy16"
style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("PK��")%>" size="60">
      </font></td>
    </tr>
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
        <input type="submit" value="�T �w">
        <font color="#CCCCCC">------- </font> 
        <input  onClick="javascript:window.document.location.href='modifygg.asp'" value="�� �s" type=button name="back">
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
