<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"asw" then
response.redirect "index.asp"
end if
   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & _
             Server.MapPath("school/message.asp")
   Set rs = Server.CreateObject("ADODB.Recordset")
   sql = "Select * from QUESTION "
   rs.Open sql, conn, adOpenStatic
%>
<script language="javascript">
function popwin(id)
{		
window.open("edit.asp?id="+id+"",'','menubar=yes,scrollbars=yes resizable=yes width=480,height=300');
}
</script>
<script language="javascript">
function click() {
if (event.button==2) {
alert('PPMM�n�ۧA�A���n���H�G���n�@����~�I�I�I')
}
}
document.onmousedown=click
</script>
<html>
<head>
<title>���D����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet>
 </head>

<body oncontextmenu=self.event.returnValue=false background="images/bg.gif" bgcolor="<%=bgcolor%>">
<form method="post" action="answer2.asp" name="form">
 <p align="center">�ٳѤU
 <input type="text" name="remain_time" disabled size="2" value="80" maxlength="2" style="width:20;height:16;border:0 solid #FFFFFF; font-size:10.8pt; ">��        

  <table width="96%" cellspacing="1" cellpadding="0" border="1" bordercolor="#000000" align="center">
    <tr> 

      <td width="18%" height="17"></td>
    </tr>
    <tr> 
      <td width="82%" height="17">
        <p style="line-height: 105%; margin-left: 6; margin-right: 6"><b>�D��</b></p>
      </td>
      <td width="18%" height="17">&nbsp;</td>
    </tr>
    <%dim i,ran
			for i=1 to 10
			randomize
			ran=(int((6-0+1)*rnd+0))&(int((9-0+1)*rnd+0))&(int((9-0+1)*rnd+0))
			ran=clng(ran)
			rs.movefirst
			rs.move ran-1
			If rs.EOF Then Exit For
			%> 
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">��<%=I%>�D�G<%=rs("question")%></p></td>
      <td width="18%">��ܵ��� 
        <input type="hidden" name="ann<%=i%>" value="<%=rs("answer")%>">
      </td>
    </tr>
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">A�G<%=rs("A")%></p></td>
      <td width="18%"> 
        <input type="radio" name="AN<%=I%>" value="A">
      </td>
    </tr>
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">B�G<%=rs("b")%></p></td>
      <td width="18%"> 
        <input type="radio" name="AN<%=I%>" value="B">
      </td>
    </tr>
    <tr> 
      <td width="82%"><p style="line-height: 105%; margin-left: 6; margin-right: 6">C�G<%=rs("c")%></p></td>
      <td width="18%"> 
        <input type="radio" name="AN<%=I%>" value="C">
      </td>
    </tr>
    <tr> 
      <td width="82%" height="26"><p style="line-height: 105%; margin-left: 6; margin-right: 6">D�G<%=rs("d")%></P></td>
      <td width="18%" height="26"> 
        <input type="radio" name="AN<%=I%>" value="D">
      </td>
    </tr>
    <%
	     Next
	    %> 
  </table>

  <div align="center"> 
    <input type="submit" name="Submit" value="����">

  </div>

</form>
</html>              <script language="javascript">        
      
function time_step()        
{        
        with(document.form.remain_time)        
        {        
                value=(value!=0)?value-1:0;        
        
                if(document.form.remain_time.value<=0)        
                {        
                        alert("�z�w�g�W�L�ɶ��A���Ѱݵ������I");        
						location.href ='INDEX.ASP'        
				}        
				else        
				{        
			       		window.setTimeout("time_step()",1000);        
				}        
		}        
}        
time_step();        
</script>         