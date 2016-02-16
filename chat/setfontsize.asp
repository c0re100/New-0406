<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(useronlinename," "&info(0)&" ")=0 then Response.Redirect "chaterr.asp?id=001"
if chatbgcolor="" then chatbgcolor="008888"%><html>
<script Language="JavaScript">
function setfs(){parent.f2.document.af.fs.value=document.forms[0].ftsz.value;parent.f2.document.af.lh.value=document.forms[0].lheight.value;this.location.href="setfontsizeok.asp";}
</script>
<head>
<title>設置字號及行距</title>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" leftmargin="0" topmargin="0">
<div align=center>
    <tr>
<td>
      <table border="1" bgcolor="E0E0E0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" cellpadding="4" width="154" height="267">
        <form method="post" action="javascript:setfs()" name="">
<tr>
            <td bgcolor="#EEFFEE"> 
              <div align=center>
<p><font color="#FF0000" style="font-size:12pt">字號及行距</font></p>
<p>請選擇你所需要的字體大小值，對話區的行距值，然後點擊“提交”</p>
<p>
<select name="ftsz" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<option value="9">9 磅</option>
<option value="9.5">9.5 磅</option>
<option value="10">10 磅</option>
<option value="10.5">10.5 磅</option>
<option value="11">11 磅</option>
<option value="11.5">11.5 磅</option>
<option value="12">12 磅</option>
<option value="12.5">12.5 磅</option>
<option value="13">13 磅</option>
<option value="13.5">13.5 磅</option>
<option value="14">14 磅</option>
<option value="14.5">14.5 磅</option>
<option value="15">15 磅</option>
<option value="15.5">15.5 磅</option>
<option value="16">16 磅</option>
</select></p>
<p>
<select name="lheight" style='font-size:12px;color:#FF6633;background-color:#EEFFEE'>
<option value="120">１２０%</option>
<option value="125">１２５%</option>
<option value="130">１３０%</option>
<option value="135">１３５%</option>
<option value="140">１４０%</option>
<option value="145">１４５%</option>
<option value="150">１５０%</option>
<option value="155">１５５%</option>
<option value="160">１６０%</option>
<option value="165">１６５%</option>
<option value="170">１７０%</option>
<option value="175">１７５%</option>
<option value="180">１８０%</option>
<option value="185">１８５%</option>
<option value="190">１９０%</option>
<option value="195">１９５%</option>
<option value="200">２００%</option>
<option value="205">２０５%</option>
<option value="210">２１０%</option>
<option value="215">２１５%</option>
<option value="220">２２０%</option>
<option value="225">２２５%</option>
<option value="230">２３０%</option>
<option value="235">２３５%</option>
<option value="240">２４０%</option>
<option value="245">２４５%</option>
</select>
</p>
<p>
<input type="submit" name="Submit" value="提交">
                </p>
</div>
</td>
</tr>
</form>
</table>
</td>
</tr>

</div>
<Script Language=Javascript>
document.forms[0].ftsz.value=parent.f2.document.af.fs.value;
document.forms[0].lheight.value=parent.f2.document.af.lh.value;
parent.m.location.href='about:blank';
</Script>
</body>
</html> 