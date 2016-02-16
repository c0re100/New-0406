<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(3)<30 then Response.Redirect "../error.asp?id=460"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 掌門 from 門派 where 掌門='"&info(0)&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經是掌門了！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select 門派,等級,銀兩,知質,會員等級 from 用戶 where id="&info(9),conn
if rs("門派")="官府" or rs("門派")="遊俠" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('好好玩，別搗亂！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("會員等級")<=1 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('２級會員以上才可以來創派！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("等級")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('掌門條件戰鬥等級大於100級！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("銀兩")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('掌門條件銀兩大於2億！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("知質")<10000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('掌門條件資質大於10000！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<html>
<head>
<title>自創門派</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="../linjianww/0064.jpg">
<form action="newmp.asp" method=POST >
  <ul>
    <table border=0 cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="49"> 
          <p align="center"><font size="2" class=c><br>
            </font><b><font size="+6" color="#FF0000">自創門派</font></b><font size="2" class=c><b><br>
            </b> <br>
            </font></p>
        </td>
      </tr>
      <tr> 
        <td height="2"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1">門派名:</font></font> 
            <input name="subject" size=27 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength=10>
            <br>
            <br>
            注：門派名稱最多為10個英文安符，5個中文字符<br>
            <br>
          </div>
        </td>
      </tr>
      <tr> 
        <td height="25"> 
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1">簡 
            介:</font></font> 
            <input name="comment" value="" size=27 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength=30>
            <br>
            <br>
            注：門派簡介最多為30個字符，且不可為空！<br>
          </div>
        </td>
      </tr>
      <tr> 
        <td height="25"> 
          <div align="center"> 
            <p><font color="#FFFFFF"><font color="#000000" size="+1"><br>
              弟子性別:</font></font> 
              <select name="ljxb" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                <option value="男">男</option>
                <option value="女">女</option>
                <option value="所有" selected>所有人</option>
              </select>
              <br>
              <br>
              注：弟子性別指該門派隻允許男/女/所有人的加入！</p>
            <p>&nbsp;</p>
          </div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="創 建">
      <input type="RESET" name="Reset" value="清 除">
      </font> </div>
  </ul>
</form>
</body>
</html>
