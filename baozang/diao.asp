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
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select wabao from ���X where id>=0 order by id",conn
wa=rs("wabao")
if wa=1 and info(4)=0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���O�|���Ц^�a�I');location.href = '../help/info.asp?type=2&name=�|���[�J��k';}</script>"
Response.End
else 
if session("wabao")="" then session("wabao")=now()
sja=session("wabao")
t=int((sja+1/1200-now())*100000)
%>
<html>
<head>
<title>�}���_��</title>
<style type="text/css">
<!--
body {  font-family: "�s�ө���"; font-size: 12px}
a:link {  font-family: "�s�ө���"; text-decoration: none}
a:visited {  font-family: "�s�ө���"; text-decoration: none}
-->
</style>
<script language=javascript>
if(window.name!="diao")
{ var i=1;
while (i<=50)
{
window.alert("�A�Q�@����r�A�§ڡH�o�̬O���檺�A�h�O�B���h�a�I���I�C�C�I50���I�I");
i=i+1;
}
top.location.href="../exit.asp"
}
</script>
</head>
<BODY BGCOLOR="#000000" link="#660000" vlink="#660000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false >
<div id="Layer1" style="position:absolute; width:200px; height:115px; z-index:1; left: 144px; top: 183px"> </div>
<div id="Layer2" style="position:absolute; width:124px; height:78px; z-index:2; left: 535px; top: 104px"><img border="0" src="wolf-admin.gif" width="75" height="50"></div>
<div align="center"><br>
  �}�l�ɶ��G<%=sja%> </div>
<div align="center">
<form name="af">
    <input type=text name='clock' style="text-align:right;font-size:12px;background-color:006600;color:FFFFFF;border: 1 double" value="" size=3 readonly>
    <script language="JavaScript"> 
function startnosay(){ 
var nosay=<%=t%>; 
document.af.clock.value=nosay; 
} 
function runnosay(){ 
document.af.clock.value=document.af.clock.value-1; 
  if(document.af.clock.value==0){ alert('�z�I�o�{���ǪF��A���֯�X�ӬݬݡI'); location.href="diao.asp"; } 
  setTimeout('runnosay()',859); 
} 
startnosay(); 
runnosay(); 
</script> 
    <a href="diao.asp">��s,�ݬݦ��S���s�o�{</a>   </form> 
  <br> 
  <% 
if sja<now()-1/1080 then 
session("wabao")="" 
%> 
�@�i�Ӥ���F��]�S�����I���I�նO�O��I<a href="diao.asp">�쨺��ݬݧa</a> 
<br>  
  <br>  
<%  
response.end  
end if  
if sja>now()-1/1200 then  
%>  
  <div id="Layer3" style="position: absolute; width: 348; height: 197; z-index: 3; left: 238; top: 199">  
<IMG SRC="wa.gif" width="135" height="78"><br> <br>  
    �H���]���A�������`�A�ګ��A�ګ��C�C�C<br>  
    <br>  
  </div>  
  <%  
else  
session("yuokaa")="yes" 
%>  
  <div id="Layer4" style="position:absolute; width:200px; height:115px; z-index:4; left: 274px; top: 197px">   
    <a href="Diaoyuok.asp"><IMG SRC="waok.gif" border="0" width="92" height="85"></a></div>  
<%  
end if  
%></div>  
</body>  
</html>
<% 
end if
rs.Close  
set rs=nothing              
conn.Close              
set conn=nothing 
 %>


