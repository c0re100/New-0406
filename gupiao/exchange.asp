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
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
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
sql="select �Ȩ� from �Τ� where �m�W='" & uname & "'"
set rs=conn.execute(sql)
yin=rs("�Ȩ�")
conn.close
set rs=nothing%>
<!--#include file="jhb.asp"-->
<%
sql="select ��� from �Ȥ� where �b��='"&uname&"'"
set rs=conn.execute(sql)
   if rs.eof or rs.bof then
%>
<script language="vbscript">
  alert "�A�٨S���Ѳ��b��A���ણ��"
  location.href = "index.asp"
</script>
<%
   else
   nowmoney=rs("���")
   set rs=nothing
   sql="select * from �Ѳ� where sid="&sid
   set rss=conn.execute(sql)
   if rss("���A")="��" then
   %>
<script language="vbscript">
  alert "���Ѫѥ��٨S�}�L�Ϊ̤w�g�ʽL�A����i����"
  location.href = "index.asp"
</script>
<%
   else
   
   sql="select ���Ѽ� from �j�� where sid="&sid&" and �b��='"&uname&"'"
   set rs=conn.execute(sql)
   if rs.eof then 
   sushare=0
   else
   sushare=rs("���Ѽ�")
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
      window.alert ("�ж�g�ʶR�ƶq!"); 
      return false;}
     for (i=0;i<form1.ushare.value.length;i++){
       n=form1.ushare.value.substr(i,1);
         if (!isnumber(n)){
            window.alert ("�п�J�@�Ӥj��0�����"); 
            return false;}
   }
return true;}
function B3_onclick(){
     if (Number(form2.ushare.value)==""){
      window.alert ("�ж�g��X�ƶq!"); 
      return false;}
     for (i=0;i<form2.ushare.value.length;i++){
       n=form2.ushare.value.substr(i,1);
         if (!isnumber(n)){
            window.alert ("�п�J�@�Ӥj��0�����"); 
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
      <p><font color="#006633">�Ѳ��N���G<%=sid%> �Ѳ��ƶq�G<%=rss("�y�q�Ѳ�")%> ��e����G<%=formatnumber(rss("��e����"),2)%></font></p>
    </center>
  </div>
  <div align="center">
    <center>
      <p><font color="#006633">�z���Ѳ��b��̲{������G<%=formatnumber(yin,0)%>��,�֦��o�تѲ��ơG<%=sushare%>,�z�i�H�R�J�ӪѲ��G<%=formatnumber((yin)/rss("��e����")-1,0)%>�� 
        </font></p>
    </center>
  </div>
  <div align="center">
    <center>
      <p><font color="#006633">�ʶR�ƶq�G
        <input name="ushare"
  size="20" class="smallinput">
        <input type="submit" value="�ʶR" name="B1"
  LANGUAGE="javascript" onclick="return B1_onclick()" class="buttonface">
        <input
  type="reset" value="����" name="B2" class="buttonface">
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
      <p><font color="#006633">��X�ƶq�G
        <input
  name="ushare" size="20" class="smallinput">
        <input type="submit" value="��X" name="B3"
  LANGUAGE="javascript" onclick="return B3_onclick()" class="buttonface">
        <input
  type="reset" value="����" name="B2" class="buttonface">
        </font></p>
    </center>
  </div>
</form>
<font color="#006633"><%end if%><%if uname="root" then%><center><a href="zhang.asp">��</a>�V�V�V�V<a href="die.asp">�^</a></center><%end if%></font>
</body>
</html>
<%
end if
conn.close
set conn=nothing
%>