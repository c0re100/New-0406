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
<html>
<head>
<title>�������</title>
<meta http-equiv=Content-Type content="text/html; charset=big5">
<LINK href="../css.css" rel=stylesheet>
</head>
<body bgcolor="#FFFFFF" background="../ff.jpg"
leftmargin="0" topmargin="7">
<div align="center">
  <script>
if (!document.layers&&!document.all)
event="test"
function showtip2(current,e,text){

if (document.all&&document.readyState=="complete"){
document.all.tooltip2.innerHTML='<marquee style="border:1px solid black">'+text+'</marquee>'
document.all.tooltip2.style.pixelLeft=event.clientX+document.body.scrollLeft+10
document.all.tooltip2.style.pixelTop=event.clientY+document.body.scrollTop+10
document.all.tooltip2.style.visibility="visible"
}

else if (document.layers){
document.tooltip2.document.nstip.document.write('<b>'+text+'</b>')
document.tooltip2.document.nstip.document.close()
document.tooltip2.document.nstip.left=0
currentscroll=setInterval("scrolltip()",100)
document.tooltip2.left=e.pageX+10
document.tooltip2.top=e.pageY+10
document.tooltip2.visibility="show"
}
}
function hidetip2(){
if (document.all)
document.all.tooltip2.style.visibility="hidden"
else if (document.layers){
clearInterval(currentscroll)
document.tooltip2.visibility="hidden"
}
}

function scrolltip(){
if (document.tooltip2.document.nstip.left>=-document.tooltip2.document.nstip.document.width)
document.tooltip2.document.nstip.left-=5
else
document.tooltip2.document.nstip.left=150
}

</script>
</div>
<div id="tooltip2" style="position:absolute;clip:rect(0 150 50 0);width:111px;background-color:#99FF99; top: 31px; left: 64px; visibility: hidden; height: 13px"> 
</div>
<div align="center">
  <table width="590" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="90" valign="top"> <P align="center"><a href="#" onClick="window.open('creat.htm','zhuangti','scrollbars=no,resizable=no,width=620,height=450')"><br>
          �۳Ъ���</a></P>
        <P align="center"><a href="#" onClick="window.open('setwg.asp','zhuangti','scrollbars=no,resizable=no,width=620,height=450')">�Z�\�]�m</a></P>
        <P ALIGN="CENTER"><a href="#" onClick="window.open('managemp.asp?id=<%=info(5)%>','zhuangti','scrollbars=no,resizable=no,width=620,height=450')">�x���޲z</a></P></td>
      <td width="554" align="center" valign="top"> <table width="500" border="1" cellpadding="1" cellspacing="0" align="left" bordercolorlight="#000000" bordercolordark="#FFFFFF">
          <tr align="center" bgcolor="#FF6600"> 
            <td width="120" height="11">�� ��</td>
            <td width="120" height="11">�x ��</td>
            <td width="80" height="11">�� �l ��</td>
            <td width="60" height="11">�A �X</td>
            <td width="100" height="11">�� �@</td>
          </tr>
          <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim jhmenpai(100,2)
rs.open "SELECT ���� FROM ����",conn
do while not rs.bof and not rs.eof
   	menpai=menpai+1
	jhmenpai(menpai,1)=trim(rs("����"))
	rs.movenext
loop
for xx=1 to menpai
tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where  ����='"&jhmenpai(xx,1)&"'")
regren=tmprs("�ƶq")
sql="update ���� set �̤l��="& regren &" where ����='"& jhmenpai(xx,1) &"'"
Set Rs=conn.Execute(sql)
next
rs.open "SELECT ����,�x��,�̤l��,�A�X,²��,�Х߮ɶ�,������� FROM ����",conn
do while not rs.bof and not rs.eof
%>
          <%if trim(rs("����"))<>"�x��" and trim(rs("����"))<>"�C�L" then%>
          <tr bgcolor="#99CCFF"> 
            <td><a href="mp.asp?my=<%=info(0)%>&amp;id=<%=rs("����")%>"><%=rs("����")%></a></td>
            <td><%=rs("�x��")%></td>
            <td><%=rs("�̤l��")%></td>
            <td><%=rs("�A�X")%></td>
            <td> <a href="#" onClick="window.open('mp1.asp?my=<%=info(0)%>&amp;id=<%=rs("����")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')" onMouseOver="showtip2(this,event,'<%=rs("²��")%> �Ӫ����إߩ�G<%=rs("�Х߮ɶ�")%> �{������G<%=rs("�������")%>')" onMouseOut="hidetip2()">�[�J</a> 
              | <a href="#" onClick="window.open('mp2.asp?my=<%=info(0)%>&amp;id=<%=rs("����")%>','zhuangti','scrollbars=no,resizable=no,width=430,height=290')">���}</a> 
            </td>
          </tr>
          <%end if%>
          <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
