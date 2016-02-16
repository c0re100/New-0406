<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>門派管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<script language="JavaScript">
function shutwin()
{
window.close();
return;
}
</script>
</head>
<body background="0064.jpg" text="#000000">
<p align="center"> <font size="+6"> 
  <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
set subj=server.createobject("adodb.recordset")
set owner=server.createobject("adodb.recordset")
subj.open "Select * From 門派",conn,0,1
if not subj.eof then
%>
  <font color="#FF3333" face="黑體">[ 門派管理 ]</font></font></p>
<p align="center"><font size="2">在門派管理裡面，長老可以更改掌門名，刪除掌門身份，和修改掌門資料！</font></p>
<table width="505" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
  <tr bgcolor="#FF6600"> 
    <td align="center" width="55"><font color="#FFFFFF" size="2">&nbsp;門 派</font></td>
    <td width="39" align="center"><font size="2" color="#FFFFFF">弟子數</font></td>
    <td width="59" align="center"><font color="#FFFFFF" size="2">招收性質</font></td>
    <td width="60" align="center"><font color="#FFFFFF" size="2">掌 門</font></td>
    <td width="103" align="center"><font size="2" color="#FFFFFF">門派基金</font></td>
    <td width="64" align="center"><font size="2" color="#FFFFFF">開除掌門</font></td>
    <td align="center" colspan="2"><font color="#FFFFFF" size="2">操 作</font></td>
  </tr>
  <%do while not subj.eof%>
  <tr bgcolor="#FFFFFF"> 
    <td width="55"> 
      <div align="center"><font size="2"><%=subj("門派")%></font></div>
    </td>
    <td width="39"> 
      <div align="right"><font size="2"><%=subj("弟子數")%>個</font></div>
    </td>
    <td width="59"> 
      <div align="center"><font size="2"><%=subj("適合")%></font></div>
    </td>
    <td width="60"><font size="2"><%=subj("掌門")%></font></td>
    <td width="103"> 
      <div align="right"><font size="2"><%=subj("幫派基金")%></font></div>
    </td>
    <td width="64"> 
      <div align="center"><font size="2"><a href=delzm.asp?username=<%=subj("門派")%> title="刪除掌門"> 
        </a><a href=delzm.asp?username=<%=subj("門派")%> title="刪除掌門">開除</a></font></div>
    </td>
    <td width="32"> 
      <p align="center"><font size="2"><a href="delmp.asp?owner=<%=SUBJ("門派")%>&action=delete" title="刪除門派">刪除</a></font></p>
    </td>
    <td width="35"> 
      <p align="center"><font size="2"><a href="managemp.asp?id=<%=SUBJ("門派")%>">修改</a></font> 
    </td>
  </tr>
  <%subj.movenext
loop %>
  <% else %>
</table>
<table width="54%" border="0" cellspacing="0" cellpadding="0" height="40" align="center">
<tr>
<td>
      <div align="center"><font size="2">刪除門派會將檔案中所有該門派的人--門派置為“無”，身份為“無”，同時刪除論壇中該門派的所有帖子，長老，考慮好再按“刪除”啊。<br>
        (刪除門派暫時還不扣掌門人的任何狀態，長老要加上也行啊)</font></div>
</td>
</tr>
</table>
<p>&nbsp;</p>
<table align="center" width="197">
<td height="14">
      <div align="center"><font size="2">還沒有一個門派呢！</font> 
        <% subj.close
set subj=nothing
end if%>
      </div>
</table>
<div align="center"></div>
<p align="center"><font size="2" color="#FF0000">[<a href="adminmpp.asp">====刷新====</a>]</font><br>
</body>
