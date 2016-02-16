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
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "請進入聊天室再進入股票！"
window.close()
</script>
<%Response.End
end if
sid=Request.QueryString ("sid")
dim uname
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select 銀兩 from 用戶 where 姓名='" & uname & "'"
set rs=conn.execute(sql)
yin=rs("銀兩")
conn.close
set rs=nothing%>
<!--#include file="jhb.asp"-->
<%
sql="select 資金 from 客戶 where 帳號='"&uname&"'"
set rs=conn.execute(sql)
   if rs.eof or rs.bof then
%>
<script language="vbscript">
  alert "你還沒有股票帳戶，不能炒股"
  location.href = "index.asp"
</script>
<%
   else
   nowmoney=rs("資金")
   set rs=nothing
   sql="select * from 股票 where sid="&sid
   set rss=conn.execute(sql)
   if rss("狀態")="封" then
   %>
<script language="vbscript">
  alert "今天股市還沒開盤或者已經封盤，不能進行交易"
  location.href = "index.asp"
</script>
<%
   else
   
   sql="select 持股數 from 大戶 where sid="&sid&" and 帳號='"&uname&"'"
   set rs=conn.execute(sql)
   if rs.eof then 
   sushare=0
   else
   sushare=rs("持股數")
   end if
%>
<script ID="clientEventHandlersJS"
LANGUAGE="javascript">
<!--
function isnumber(c){
 	if ((c>='0') && (c<='9')) 
		return true; 
	else 		
        return false;
     } 

function B1_onclick(){
     if (Number(form1.ushare.value)==""){
      window.alert ("請填寫購買數量!"); 
      return false;}
     for (i=0;i<form1.ushare.value.length;i++){
       n=form1.ushare.value.substr(i,1);
         if (!isnumber(n)){
            window.alert ("請輸入一個大於0的整數"); 
            return false;}
   }
return true;}
function B3_onclick(){
     if (Number(form2.ushare.value)==""){
      window.alert ("請填寫賣出數量!"); 
      return false;}
     for (i=0;i<form2.ushare.value.length;i++){
       n=form2.ushare.value.substr(i,1);
         if (!isnumber(n)){
            window.alert ("請輸入一個大於0的整數"); 
            return false;}
   }
return true;}
//-->
</script>
<html>

<head>
<title></title>
<link rel="stylesheet" href="../../CSS/STYLE.CSS">
</head>

<body bgcolor="#FFFEF4">
<form name="form1" method="post" action="buy.asp">
  <input type="hidden" name="sid" value="<%=sid%>">
  <div align="center">
    <center>
      <p><font color="#006633">股票代號：<%=sid%> 股票數量：<%=rss("流通股票")%> 當前價格：<%=formatnumber(rss("當前價格"),2)%></font></p>
    </center>
  </div>
  <div align="center">
    <center>
      <p><font color="#006633">您的股票帳戶裡現有資金：<%=formatnumber(yin,0)%>元,擁有這種股票數：<%=sushare%>,您可以買入該股票：<%=formatnumber((yin)/rss("當前價格")-1,0)%>股 
        </font></p>
    </center>
  </div>
  <div align="center">
    <center>
      <p><font color="#006633">購買數量：
        <input name="ushare"
  size="20" class="smallinput">
        <input type="submit" value="購買" name="B1"
  LANGUAGE="javascript" onclick="return B1_onclick()" class="buttonface">
        <input
  type="reset" value="取消" name="B2" class="buttonface">
        </font></p>
    </center>
  </div>
</form>
<form name="form2" method="post" action="sale.asp">
  <font color="#006633">
  <input type="hidden" name="sid" value="<%=sid%>">
  </font>
  <div align="center">
    <center>
      <p><font color="#006633">賣出數量：
        <input
  name="ushare" size="20" class="smallinput">
        <input type="submit" value="賣出" name="B3"
  LANGUAGE="javascript" onclick="return B3_onclick()" class="buttonface">
        <input
  type="reset" value="取消" name="B2" class="buttonface">
        </font></p>
    </center>
  </div>
</form>
<font color="#006633"><%end if%><%if uname="root" then%><center><a href="zhang.asp">漲</a>––––<a href="die.asp">跌</a></center><%end if%></font>
</body>
</html>
<%
end if
conn.close
set conn=nothing
%>