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
grade=info(2)
myid=info(9)
myid=int(Request.QueryString("id"))
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>琩珇</title>
<style>
td{font-size:9pt}
</style>

<script  language="javascript">  
 
function  check() 
{
if (jiaoyi.sl.value>jiaoyi.sl2.value) {
alert("扳计秖禬局Τ计秖");
jiaoyi.sl.focus(); 
return  false;}

if  (jiaoyi.money.value==""){
alert("﹚基ぃ瓳");
jiaoyi.money.focus(); 
return  false;}

if(isNaN(jiaoyi.money.value)) {
alert("﹚基唉块计");
jiaoyi.money.focus(); 
return  false;}
if (jiaoyi.money.value <0 ) {
alert("﹚基基ぃ0");
jiaoyi.money.focus(); 
return  false;}
if (jiaoyi.money.value >1000000000 ) {
alert("﹚基基ぃ10货");
jiaoyi.money.focus(); 
return  false;}
return  true  
}
</script>  


</head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed">
<p style="font-size:12pt; color:red" align="center"><br>
  <font color="#FFFFFF">ユ珇</font></p>
<div align="center"></div>

<form method=POST action='jiaoyi1.asp' name="jiaoyi"  onsubmit="return  check()">
  <table border="1" cellspacing="0" cellpadding="4" width="462" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" align="center" >
    <tr> 
    <td width="13%" > 
      <div align="center"> 珇 </div>
    </td>
    <td width="13%" > 
      <div align="center">瞷Τ计秖</div>
    </td>
    <td width="15%" >ユ计秖</td>
    <td width="30%" > 
      <div align="center">﹚基</div>
    </td>
    <td width="29%" > 
      <div align="center">ㄏノ</div>
    </td>
  </tr>
  <%
  

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 珇,计秖,摸 FROM 珇 WHERE 局Τ="&info(9)&" and 计秖>=1 and 杆称<>true and id="&myid
rs.open sql,conn,1,1
do while not rs.bof and not rs.eof
shuliang=rs("计秖")
%>
  <tr > 
    <td width="13%" valign="middle" height="18" ><%=rs("珇")%> <font color=#ffffff> 
        <input type=hidden name=wupin value="<%=rs("珇")%>">
      </font></td>
    <td width="13%" height="18"> 
      <center>
          <%=rs("计秖")%><font color=#ffffff> </font><font color=#ffffff>
          <input type=hidden name=sl2 value="<%=rs("计秖")%>">
          </font> <font color=#ffffff> </font> 
        </center>
    </td>
    <td width="15%" height="18"> 
      <select name=sl size=1>
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
<option value="7">7</option>
<option value="8">8</option>
<option value="9">9</option>

      </select>
    </td>
      <td width="30%" height="18"><font color=#ffffff> 
        <input 
      style="BORDER-RIGHT: 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: 1px solid; BORDER-BOTTOM: 1px solid" 
      maxlength=10 size=10 name=money>
        <select size=1 name=danwei>
          <option value="蝗ㄢ">蝗ㄢ</option>
          <option value="刽">刽</option>
        </select>
        </font></td>
    <td width="29%" height="18">
      <input name=ok type=submit value=扳>
    </td>
  </tr>
  <%
rs.movenext
loop

rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
</form>
</body>
</html>