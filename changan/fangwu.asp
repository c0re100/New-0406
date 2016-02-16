<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
name=info(0)%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM 房屋 where  類型='一般房屋'"
Set Rs=conn.Execute(sql)
%>
<html>
<link href="../chat/ccss2.css" rel="stylesheet" type="text/css">

<body bgcolor="#FFFDDF" background="../ff.jpg">
<table border=0 bgcolor="#FFFFCC" align=center width=98% cellpadding="2" cellspacing="2" bordercolor="#65251C">
  <tr> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099"> 
        <a href="../jh.asp">返回江湖</a> 歡迎</font><%=name%><font color="#990099">光臨房產中心，誰都想有自己的家，這裡可以滿足您的各種要求：(現在賣出房子隻能得到原來的1/5的價格)</font></marquee></font></p>
  <tr> 
    <td height="21" colspan="7" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2" bordercolor="#6633FF">
        <tr bgcolor="#FF0000"> 
          <td> 
            <div align="center"><a href=fangwu.asp><font color="#FFFFFF">一般房屋</font></a></div>
          </td>
          <td> 
            <div align="center"><a href=fangwu3.asp><font color="#FFFFFF">高級公寓</font></a></div>
          </td>
          <td> 
            <div align="center"><a href=fangwu4.asp><font color="#FFFFFF">花園洋房</font></a></div>
          </td>
          <td> 
            <div align="center"><a href=fangwu5.asp><font color="#FFFFFF">豪華別墅</font></a></div>
          </td>
        </tr>
      </table>
  <tr bordercolor="#000000" bgcolor="#FF0000"> 
    <td height="15" width="14%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">房屋類型</font></div>
    </td>
    <td height="15" width="9%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">售價</font></div>
    </td>
    <td height="15" width="11%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">城市</font></div>
    </td>
    <td height="15" colspan="2" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">門牌號碼</font></div>
    </td>
    <td height="15" width="19%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">房主/伴侶</font></div>
    </td>
    <td height="15" width="25%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">操作</font></div>
    </td>
  </tr>
  <% 
do while not rs.eof and not rs.bof 
%> 
  <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="14%" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%=rs("類型")%> 
      </center>
      </font> 
    <td width="9%" class="calen-curr" height="15"><%=rs("售價")%> 
      <div align="center"></div>
    <td width="11%" class="ct-def4" height="15"> 
      <div align="center"><font size="2"><%=rs("位置")%></font></div>
    </td>
    <td colspan="2" class="ct-def4" height="15"><%=rs("街道")%> <font color=red><%=rs("id")%></font><font size="2">號</font></td>
    <td width="19%" class="calen-curr" height="15"> 
      <div align="center"><font size="2"><%=rs("戶主")%> / </font><font color="ff00ff"><%=rs("伴侶")%></font></div>
    </td>
    <td width="25%" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%if rs("戶主")<>name and rs("戶主")="無" then%>
        <a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">購買</font></a>
        <%end if%>
        <%if rs("戶主")=name then%>
        <a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>賣出</font></a>
        <%end if%>
        <%if rs("戶主")<>"無" and rs("戶主")<>name then%>
        該房己出售
        <%end if%>
      </center>
      </font></td>
  </tr>
  <% 
rs.movenext 
loop 

	set rs=nothing	
	conn.close
	set conn=nothing
%> 
</table>
</html>