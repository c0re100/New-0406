<%@ LANGUAGE=VBScript codepage ="950" %>
<!--#include file="dbpath.asp"-->
<html>
<head>
<title>網吧聯盟</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type="text/css">
<!--
td           { font-size: 9pt; line-height: 13pt}
a:link       { color: ; text-decoration: none }
a:visited    { color: ; text-decoration: none }
a:hover      { color: #FF0000; text-decoration: none }
-->
</style>
</head>
<body topmargin="6">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr>
    <td width="135" valign="top">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td width="100%"><map name="FPMap0">
      <area href="./" coords="2, 0, 134, 46" shape="rect"></map><img border="0" src="img/a1.gif" usemap="#FPMap0"></td>
        </tr>
        <tr>
          <td width="100%" height="5"></td>
        </tr>
        <tr>
          <td width="100%">
      <table border="0" width="100%" cellspacing="1" cellpadding="0" background="img/l.gif">
        <tr>
          <td width="100%"><form method="post" action="index.asp?type=se">
                    <div align="center">
                      <center>
                    <table width="123" border="0" cellspacing="0" cellpadding="0" class="base" height="92">
                      <tr background="images/l.gif"> 
                        <td align="center" height="20">所在城市 </td>
                      </tr>
                      <tr> 
                        <td align="center" height="38"> 
                        <input type="text" name="city" size="10">
                        </td>
                      </tr>
                      <tr background="images/l.gif"> 
                        <td align="center" height="18">接入方式 </td>
                      </tr>
                      <tr> 
                        <td align="center" height="37"> 
                          <select name="typical" size="1">
                            <option selected value="">所有方式</option>
                            <option value="電話接入">電話接入</option>
                            <option value="ISDN接入">ISDN接入</option>
                            <option value="DDN接入">DDN接入</option>
                            <option value="ADSL接入">ADSL接入</option>
                            <option value="其它方式">其它方式</option>
                          </select>
                        </td>
                      </tr>
                      <tr background="images/l.gif"> 
                        <td align="center" height="20">網吧名稱 </td>
                      </tr>
                      <tr> 
                        <td align="center" height="28"> 
                          <input type="text" name="barname" size="10">
                        </td>
                      </tr>
                      <tr background="images/l.gif"> 
                        <td align="center" height="20">所在地址 </td>
                      </tr>
                      <tr> 
                        <td align="center" height="28"> 
                          <input type="text" name="add" size="10">
                        </td>
                      </tr>
                      <tr> 
                        <td align="center" height="2"> 
                          <input type="submit" value="查詢">
                        </td>
                      </tr>
                    </table>
                      </center>
                    </div>
                  </form></td>
        </tr>
      </table>
</td>
        </tr>
      </table>
    </td>
    <td width="567" valign="top">
      <table border="0" width="100%" cellspacing="0" cellpadding="0">
        <tr>
          <td width="22%" valign="top"><img border="0" src="img/a2.gif"></td>
          <td width="78%">
            <p align="center"><a href="wbjm.asp">加盟條件</a>&nbsp;<a href="index.asp?type=reg">我要加盟</a>&nbsp;<a href="index.asp?type=login">資料修改</a>&nbsp;<a href="index.asp?type=top">人氣排行</a>&nbsp;</td>
        </tr>
        <tr>
          <td width="22%" valign="top"></td>
          <td width="78%">
            <hr size="1">
          </td>
        </tr>

      </table> 
      <div align="right"> 
        <table border="0" width="490" cellspacing="0" cellpadding="0"> 
          <tr> 
            <td width="100%" height="70" valign="middle"> 
              <p align="center"><a href="index.asp"><img border="0" src="img/a4.gif"></a><img border="0" src="img/a5.gif"></td> 
          </tr> 
        </table> 
      </div> 
     <%if request("type")="reg"  then%> 
      <div align="right"> 
        <table border="0" width="490" cellspacing="0" cellpadding="0"> 
          <tr> 
            <td width="100%"> 
              <p align="center"><font color="#ff0000">請您認真填寫以下項目，我們將對您所輸入的信息進行檢查，<br> 
              以便確認是否有不符合規定的地方！</font></td> 
          </tr> 
        </table> 
      </div> 
     <%end if%> 
      <table border="0" width="100%" cellspacing="5" cellpadding="0"> 
        <tr> 
          <td width="100%"></td> 
        </tr> 
      </table> 
  <%if request("type")="reg"   then%> 
<script language=JavaScript> 
<!-- 
function checkvalue() 
 { 
 if(form.barname.value.length>3) 
  { 
  } 
  else 
  { 
  alert("請輸網吧名稱，至少四位！"); 
  return false; 
  } 
 if(form.people.value.length>2) 
  { 
  } 
  else 
  { 
  alert("請輸入聯繫人姓名，至少四位！"); 
  return false; 
  } 
 if(form.add.value.length>5) 
  { 
  } 
  else 
  { 
  alert("請輸入網吧地址，至少六位！"); 
  return false; 
  } 
 if(form.intros.value.length>20) 
  { 
  } 
  else 
  { 
  alert("請輸入網吧簡介，最少10個漢字！"); 
  return false; 
  } 
  if(document.form.pwd.value!="") 
  { 
  } 
  else 
  { 
  alert("請輸入密碼！"); 
  return false; 
  } 
   
 if(document.form.pwd.value==document.form.pwd2.value) 
  { 
  } 
  else 
  { 
  alert("兩次輸入的密碼不一致！"); 
  return false; 
  } 
 
 return true 
 } 
--> 
</script>   
 
      <div align="right"> 
        <table border="0" width="490" cellspacing="0" cellpadding="0"> 
          <tr> 
            <td width="100%"><form action="regok.asp" method="post" name="form" onsubmit="return checkvalue()"> 
  <table align="center" bgColor="#333333" border="0" cellSpacing="1" class="base" width="425"> 
    <tbody> 
      <tr bgColor="#ffffff"> 
        <td align="right" width="86">所在城市：</td> 
        <td width="329">   <input type="text" name="city" size="10"></td> 
      </tr> 
      <tr bgColor="#ffffff"> 
        <td align="right" width="86">網吧名稱：</td> 
        <td width="329"><input name="barname" size="20"> <font color="#ff0000">*</font></td>                                                                                                                              
      </tr>                                                                                                                              
      <tr bgColor="#ffffff">                                                                                                                              
        <td align="right" width="86">接入方式：</td>                                                                                                                              
        <td width="329"><select name="typical">                                                                                                                              
            <option selected value="電話接入">電話接入</option>
            <option value="ISDN接入">ISDN接入</option>
            <option value="DDN接入">DDN接入</option>
            <option value="ADSL接入">ADSL接入</option>
            <option value="其它方式">其它方式</option>
          </select></td>
      </tr>
      <tr bgColor="#ffffff">
        <td align="right" width="86">機器臺數：</td>
        <td width="329"><input name="number" size="10"> 臺</td>
      </tr>
      <tr bgColor="#ffffff">
        <td align="right" width="86">網吧地址：</td>
        <td width="329"><input name="add" size="20"> <font color="#ff0000">*</font></td>                                                                                                                              
      </tr>                                                                                                                              
      <tr bgColor="#ffffff">                                                                                                                              
        <td align="right" width="86">聯繫人：</td>                                                                                                                              
        <td width="329"><input name="people" size="20"> <font color="#ff0000">*</font></td>                                                                                                                              
      </tr>                                                                                                                              
      <tr bgColor="#ffffff">                                                                                                                              
        <td align="right" width="86">密 碼：</td>
        <td width="329"><input name="pwd" type="password" value size="20"> <font color="#ff0000">*&nbsp;</font>                                                                                                    
          不要有空格</td>                                                                                                                             
      </tr>                                                                                                                             
      <tr bgColor="#ffffff">                                                                                                                             
        <td align="right" width="86">重復密碼：</td>                                                                                                                             
        <td width="329"><input name="pwd2" type="password" value size="20"></td>                                                                                                                             
      </tr>                                                                                                                             
      <tr bgColor="#ffffff">                                                                                                                             
        <td align="right" width="86">聯繫電話：</td>                                                                                                                             
        <td width="329"><input name="phone" size="20"></td>                                                                                                                             
      </tr>                                                                                                                             
      <tr bgColor="#ffffff">                                                                                                                             
        <td align="right" width="86">Email：</td>                                                                                                                            
        <td width="329"><input name="email" size="20"></td>                                                                                                                            
      </tr>                                                                                                                            
      <tr bgColor="#ffffff">                                                                                                                            
        <td align="right" width="86">郵政編碼：</td>                                                                                                                            
        <td width="329"><input name="zip" size="20"></td>                                                                                                                            
      </tr>                                                                                                                            
      <tr bgColor="#ffffff">                                                                                                                            
        <td align="right" width="86">網吧IP：</td>                                                                                                                            
        <td width="329"><input name="ip" size="20"><font color="#ff0000">*&nbsp;</font>以此ip為準！</td>                                                                                                                            
      </tr>                                                                                                                            
      <tr bgColor="#ffffff">                                                                                                                            
        <td align="right" width="86">網吧主頁：</td>                                                                                                                            
        <td width="329"><input name="homepage" size="20"></td>                                                                                                                            
      </tr>                                                                                                                            
      <tr bgColor="#ffffff">                                                                                                                            
        <td align="right" vAlign="top" width="86">網吧簡介：</td>                                                                                                                            
        <td width="329"><textarea cols="40" name="intros" rows="6"></textarea> <font color="#ff0000">*</font></td>                                                                                                                              
      </tr>                                                                                                                              
      <tr bgColor="#ffffff">                                                                                                                              
        <td align="middle" colSpan="2"><input name="UserIP" type="hidden"><input name="Submit2" type="submit" value="登記">                                                                                                                               
          <input type="reset" value="重來" name="B2"></td>                                                                                                                             
      </tr>                                                                                                                             
    </tbody>                                                                                                                             
  </table>                                                                                                                             
</form></td>                                                                                                                             
          </tr>                                                                                                                             
        </table>                                                                                                                             
      </div>                                                                                                  
<%elseif request("type")="regok" then%>                                                                                                  
      <div align="right">                                                                                                 
        <table border="0" width="490" cellspacing="0" cellpadding="0" height="60%">                                                                                                 
  <tr>                                                                                                 
    <td width="100%">                                                                                                 
      <p align="center">恭喜，您的資料已經錄入！謝謝你的支持！<br>                                                                                                 
      <a href="./">[返回首頁]</a></p>                                                                                                 
    </td>                                                                                                 
  </tr>                                                                                                 
</table>                                                                                                 
      </div>                                                                                                 
<%elseif request("type")="login" then%>                                                                                                                               
     <div align="right">                                                                                                                            
        <table border="0" width="490" cellspacing="0" cellpadding="0" height="60%">                                                                                                                            
          <tr>                                                                                                                            
            <td width="100%"><form action="user.asp?type=login" method="post" name="form2">                                                                                                                           
  <table align="center" bgColor="#333333" border="0" cellSpacing="1" class="base" width="188">                                                                                                                           
    <tbody>                                                                                                                           
      <tr bgColor="#ffffff">                                                                                                                           
        <td width="56">網吧名稱</td>                                                                                                                           
        <td width="125"><input name="barname" size="12" value="<%=request("barname")%>"></td>                                                                                                                           
      </tr>                                                                                                                           
      <tr bgColor="#ffffff">                                                                                                                           
        <td width="56">密碼</td>                                                                                                                           
        <td width="125"><input name="pwd" size="12" type="password" value></td>                                                                                                                           
      </tr>                                                                                                                           
      <tr align="middle" bgColor="#ffffff">                                                                                                                           
        <td colSpan="2"><input name="Submit2" type="submit" value="登錄"></td>                           
      </tr>                           
    </tbody>                           
  </table>                           
</form></td>                           
          </tr>                           
        </table>                           
      </div>                             
<%elseif request("type")="admin"  then%>                         
<%sqladm="select * from admin"                         
Set rsadm= Server.CreateObject("ADODB.Recordset")                         
rsadm.open sqladm,conn,1,2%>                         
      <div align="right">                            
<table border="0" width="490" cellspacing="0" cellpadding="0">                            
 <tr>                       
  <form action="user.asp?type=pwd" method="post">                            
    <td width="100%"><b>修改密碼：</b> 用戶名：<input name="user" value="<%=rsadm("user")%>" size="13">                       
      密碼：<input name="pwd" type="password" value="<%=rsadm("pwd")%>" size="13">  <input name="Submit2" type="submit" value="修改">                                                                                                                              
    </td>                           
  </form>                       
  </tr>                           
 <tr>                       
    <td width="100%">                       
      <hr color="#666666" size="1">                      
    </td>                          
 </tr>                      
</table>                       
</div>                     
<%rsu.close                                           
set rsu=nothing%>                     
                     
<%elseif request("show")<>"" then%>                                                                                               
<%sqlu="select * from bar where barname='"&request("show")&"'"                     
Set rsu= Server.CreateObject("ADODB.Recordset")                     
rsu.open sqlu,conn,1,2                     
if session("jybarcount")="" then                     
session("jybarcount")="yyy"                     
	rsu("count").value = rsu("count").value + 1                     
	rsu.Update()                                                                
end if%>                                                                
      <div align="right">                                                                                                   
        <table border="0" width="490" cellspacing="0" cellpadding="0">                                                                                                   
          <tr>                                                                                                   
            <td width="60%">                                                                                                   
              <table align="center" bgColor="#666666" border="0" cellPadding="3" cellSpacing="1" class="base" height="272" width="464">
                <form action="user.asp?type=edit&barname=<%=rsu("barname")%>" method="post">
                  <tr bgColor="#ffffff"> 
                    <td colSpan="4" nowrap> 
                      <div align="center"> <font color="#990000">網吧詳細資料</font> 
                        <%if session("barlogin")=rsu("barname") and ljjh_grade<>10 then%>
                        &nbsp;&nbsp;&nbsp;&nbsp; [<a href="user.asp?barname=<%=rsu("barname")%>&type=exit"><font color="#FF0000">退出登陸</font></a>] 
                        <%end if%>
                      </div>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td height="21" width="67" nowrap>所在城市：</td>
                    <td height="21" width="164"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input type="text" name="city" size="10" value="<%=rsu("city")%>">
                      <%else%>
                      <%=rsu("city")%> 
                      <%end if%>
                    </td>
                    <td height="21" width="61" nowrap>網吧名稱：</td>
                    <td height="21" width="133"><%=rsu("barname")%></td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td height="29" width="67" nowrap>接入類型：</td>
                    <td height="29" width="164"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <select name="typical" size="1">
                        <option value="電話接入">電話接入</option>
                        <option value="ISDN接入">ISDN接入</option>
                        <option value="DDN接入">DDN接入</option>
                        <option value="ADSL接入">ADSL接入</option>
                        <option value="其它方式">其它方式</option>
                        <option selected value="<%=rsu("typical")%>"><%=rsu("typical")%></option>
                      </select>
                      <%else%>
                      <%=rsu("typical")%> 
                      <%end if%>
                    </td>
                    <td height="29" width="61" nowrap>機器臺數：</td>
                    <td height="29" width="133"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="number" size="12" value="<%=rsu("number")%>">
                      <%else%>
                      <%=rsu("number")%> 
                      <%end if%>
                      臺</td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td width="67" nowrap>網吧地址：</td>
                    <td colSpan="3"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="add" size="52" value="<%=rsu("add")%>">
                      <%else%>
                      <%=rsu("add")%> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td width="67" nowrap>主機IP：</td>
                    <td colspan="3"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="ip" size="22" value="<%=rsu("ip")%>">
                      <%else%>
                      保密 
                      <%end if%>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td width="67" nowrap>管理員：</td>
                    <td width="164"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="people" size="12" value="<%=rsu("people")%>">
                      <%else%>
                      <%=rsu("people")%> 
                      <%end if%>
                    </td>
                    <td width="61" nowrap>電話：</td>
                    <td width="133"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="phone" size="17" value="<%=rsu("phone")%>">
                      <%else%>
                      <%=rsu("phone")%> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td width="67" nowrap>郵政編碼：</td>
                    <td width="164"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="zip" size="12" value="<%=rsu("zip")%>">
                      <%else%>
                      <%=rsu("zip")%> 
                      <%end if%>
                    </td>
                    <td width="61" nowrap>Email：</td>
                    <td width="133"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="email" size="17" value="<%=rsu("email")%>">
                      <%else%>
                      <a href="mailto:<%=rsu("email")%>"><%=rsu("email")%></a> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td width="67" nowrap>網吧主頁：</td>
                    <td colSpan="3"> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <input name="homepage" size="52" value="<%=rsu("homepage")%>">
                      <%else%>
                      <a href="<%=rsu("homepage")%>" target="_blank"><%=rsu("homepage")%></a> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td colSpan="2" nowrap>網吧簡介： 注冊日期：<%=rsu("date")%></td>
                    <td colSpan="2" nowrap>人氣指數：<font color="#ff0000"><%=rsu("count")%></font></td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td colSpan="4" nowrap> 
                      <%if session("barlogin")=rsu("barname") then%>
                      <textarea rows="2" name="intros" cols="60"><%=rsu("intros")%></textarea>
                      <%else%>
                      <font color="#333333"><%=rsu("intros")%></font> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr bgColor="#ffffff"> 
                    <td colSpan="2" height="7" nowrap> 
                      <div align="center"> 
                        <%if session("barlogin")=rsu("barname") then%>
                        <input name="Submit2" type="submit" value="修改">
                        <input type="reset" value="重置" name="B2">
                        <%else%>
                        <font color="#990000"><a href="index.asp?type=login&barname=<%=rsu("barname")%>">在線修改</a></font> 
                        <%end if%>
                      </div>
                    </td>
                    <td colSpan="2" height="7" nowrap> 
                      <div align="center"> 
                        <%if session("barlogin")=rsu("barname") then%>
                        密碼： 
                        <input name="pwd" size="12" value="<%=rsu("pwd")%>" type="password">
                        <%else%>
                        <font color="#990000">關閉窗口</font> 
                        <%end if%>
                      </div>
                    </td>
                  </tr>
                </form>
              </table>
            </td>                                       
          </tr>                                       
        </table>                                       
      </div>                                    
<%else%>                                    
<%                                 
   if not isempty(request("page")) then                                 
      currentPage=cint(request("page"))                                 
   else                                 
      currentPage=1                                 
   end if                                 
                                 
   MaxPerPage=15 '###每頁顯示條數                                 
                                 
   dim totalPut                                 
   dim CurrentPage                                 
   dim TotalPages                                 
   dim i,j                                 
                            
if request("type")="se" then
if request.form("city")="" or request.form("city")="no" then
  city=request("city")
else
  city=request.form("city")
end if
if request.form("typical")="" or request.form("typical")="no" then
  typical=request("typical")
else
  typical=request.form("typical")
end if
if request.form("barname")="" then
  barname=request("barname")
else
  barname=request.form("barname")
end if
if request.form("add")="" then
  add=request("add")
else
  add=request.form("add")
end if
sqlu="select * from bar where city like '%"&city&"%' and typical like '%"&typical&"%' and barname like '%"&barname&"%' and add like '%"&add&"%' order by id desc"
elseif request("type")="top" then                                
sqlu="select * from bar order by count desc"                              
else                                
sqlu="select * from bar order by id desc"                            
end if                                
                                
Set rsu= Server.CreateObject("ADODB.Recordset")                                
rsu.open sqlu,conn,1,1                                
%>                                
<%                                
rsu.pagesize=MaxPerPage                                
totalPut=rsu.recordcount                                                                    
        if currentpage<1 then                                                                    
          currentpage=1                                                                    
      end if                                                                    
      if (currentpage-1)*MaxPerPage>totalput then                                                                    
	   if (totalPut mod MaxPerPage)=0 then                                                                    
	     currentpage= totalPut \ MaxPerPage                                                                    
	   else                                                                    
	      currentpage= totalPut \ MaxPerPage + 1                                                                    
	   end if                                                                    
      end if                                                                    
       if currentPage=1 then                                                                    
           showpage totalput,MaxPerPage,"search.asp"                                                                    
            showContent                                                                    
            showpage totalput,MaxPerPage,"search.asp"                                                                    
       else                                                                    
          if (currentPage-1)*MaxPerPage<totalPut then                                                                    
            rsu.move  (currentPage-1)*MaxPerPage                                                                    
            dim bookmark                                                                    
            bookmark=rsu.bookmark                                                                    
           showpage totalput,MaxPerPage,"search.asp"                                                                    
            showContent                                                                    
             showpage totalput,MaxPerPage,"search.asp"                                                                    
        else                                                                    
	        currentPage=1                                                                    
           showpage totalput,MaxPerPage,"search.asp"                                                                    
           showContent                                                                    
           showpage totalput,MaxPerPage,"search.asp"                                                                    
	      end if                                                                    
	   end if                                                                    
	sub showContent                                                                    
       dim i                                                                    
       i=0                                                                    
end sub                                               
                                                                                              
  dim n                                                                         
  if totalnumber mod maxperpage=0 then                                                                         
     n= totalnumber \ maxperpage                                                                         
  else                                                                         
     n= totalnumber \ maxperpage+1                                                                         
  end if                                                                                                                                                           
%>                                
<div align="right">                                
        <table border="0" width="514" cellspacing="0" cellpadding="0">
          <tr>                                
    <td width="100%">                                
              <div align="center"> 
                <table border="0" width="501" cellspacing="1" bgcolor="#333333">
                  <tr> 
                    <td bgcolor="#FFFFFF" colspan="6"> 
                      <p align="center">網 吧 聯 盟 成 員 名 單</p>
                    </td>
                  </tr>
                  <tr> 
                    <td width="50" bgcolor="#FFFFFF" align="center"><a href="index.asp?type=id"><u> 
                      <%if request("type")="id" then%>
                      <%else%>
                      <font color="#000000"> 
                      <%end if%>
                      編號</font></u></a></td>
                    <td width="76" bgcolor="#FFFFFF" align="center">網吧名稱</td>
                    <td width="213" bgcolor="#FFFFFF" align="center">網吧地址</td>
                    <td width="34" bgcolor="#FFFFFF" align="center"><a href="index.asp?type=top"><u> 
                      <%if request("type")="top" then%>
                      <%else%>
                      <font color="#000000"> 
                      <%end if%>
                      人氣</font></u></a></td>
                    <td width="44" bgcolor="#FFFFFF" align="center">會員</td>
                    <td width="65" bgcolor="#FFFFFF" align="center">日期</td>
                  </tr>
                  <%if rsu.eof and rsu.bof then%>
                  <tr> 
                    <td bgcolor="#FFFFFF" colspan="6" height="50" align="center"> 
                      <p align="center">沒有或沒有找到任何網吧！</p>
                    </td>
                  </tr>
                  <%else%>
                  <%do while not rsu.eof%>
                  <tr> 
                    <td width="50" bgcolor="#FFFFFF" align="center" nowrap><%=rsu("id")%> 
                      
                    </td>
                    <td width="76" bgcolor="#FFFFFF" align="center"> 
                      
                      </a><a href="index.asp?show=<%=rsu("barname")%>"> 
                      
                      <%=rsu("barname")%></a></td>
                    <td width="213" bgcolor="#FFFFFF" align="center"><%=rsu("add")%></td>
                    <td width="34" bgcolor="#FFFFFF" align="center" nowrap><%=rsu("count")%></td>
                    <td width="44" bgcolor="#FFFFFF" align="center">
                      <%if rsu("qdlm")=true then%><img src="hy.gif" width="16" height="18"><%end if%>
                      </td>
                    <td width="65" bgcolor="#FFFFFF" align="center"><%=rsu("date")%></td>
                  </tr>
                  <% i=i+1                                    
 if i>=MaxPerPage then exit do                                    
 rsu.movenext                                    
loop                                    
%>
                  <tr> 
                    <td colspan="3" bgcolor="#FFFFFF"> 
                      <p> 
                        <%if request("type")="se" then%>
                        找到 
                        <%else%>
                        共有 
                        <%end if%>
                        <font color="#FF0000"><%=rsu.recordcount%></font>網吧</p>
                    </td>
                    <td colspan="3" bgcolor="#FFFFFF"> 
                      <%                                   
   mpage=rsu.pagecount     '得到總頁數                                     
   page=request("page")                                  
   if page="" then                                  
   page="1"                                  
   end if                                  
                                     
   if currentPage > 1 then                                  
      Response.Write "<A HREF=index.asp?type="&request("type")&"&city="&city&"&typical="&typical&"&barname="&barname&"&add="&add&">首頁</A> "                                       
      Response.Write "<A HREF=index.asp?type="&request("type")&"&city="&city&"&typical="&typical&"&barname="&barname&"&add="&add&"&Page=" & (page-1) & ">上一頁</A> "                                  
   end if                                  
   if currentPage < mpage then                                  
      Response.Write "<A HREF=index.asp?type="&request("type")&"&city="&city&"&typical="&typical&"&barname="&barname&"&add="&add&"&Page=" & (page+1) & ">下一頁</A> "                                       
      Response.Write "<A HREF=index.asp?type="&request("type")&"&city="&city&"&typical="&typical&"&barname="&barname&"&add="&add&"&Page=" & mpage & ">尾頁</A> "                                       
   end if                                  
%>
                      <div align="right">第<font color="#FF0000"><%=page%>/<%=mpage%></font>頁</div>
                    </td>
                  </tr>
                  <%end if%>
                </table>                                      
      </div>                                      
    </td>                                      
  </tr>                                      
</table>                                      
      </div>                                      
<%end if%>                                         
    </td>                                           
    <td width="56" valign="top">                                           
      <p align="right"><img border="0" src="img/a3.gif"></td>                                           
  </tr>                                           
</table>                                           
<table border="0" width="100%" cellpadding="0">                                          
  <tr>                                          
    <td width="100%">                                          
      <hr color="#800000">                                          
    </td>                                          
  </tr>                                          
  <tr>                                          
                                            
  </tr>                                          
</table>                                          
</body>                                          
<%rsu.close                      
set rsu=nothing                      
conn.close                      
set conn=nothing%>                      
</html>