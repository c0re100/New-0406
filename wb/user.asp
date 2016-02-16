<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%><%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")


%>
<!--#include file="dbpath.asp"-->
<%if request("type")="exit" then
session("barlogin")=""
response.redirect "index.asp?show="&request("barname")&""%>

<%elseif request("type")="adminexit" then
response.redirect "index.asp"%>

<%'######################################
elseif request("type")="login" then
    if request.form("barname")="" then
       response.write("<script>alert('網吧名稱不能為空！');history.go(-1)</script>")
    end if%>
<%set rs=server.createobject("adodb.recordset")
 sql="select * from bar where barname='"&request.form("barname")&"'"
 rs.open sql,conn,1,1
  if request.form("pwd")=rs("pwd") and rs("pwd")<>"" then
      session("barlogin")=request.form("barname")
  else
  response.write("<script>alert('密碼不對！');history.go(-1)</script>")
  end if
rs.close  
set rs=nothing
response.redirect "index.asp?show="&request.form("barname")&""
%>
<%'######################################
elseif request("type")="del" then
    
    if request("id")="" then
       response.write("<script>alert('參數錯誤:編號不能為空！')</script>")
    end if%>
<%set rst=server.CreateObject("ADODB.RecordSet")
			conn.execute "delete * from bar where id="+request("id")
			rst.close
%>
■■■■■■■■■■<br>■■■■■■■■■■<br>刪除完成！
<script LANGUAGE="JavaScript">
<!--
setTimeout('window.close();', 2000);
// -->
</script>
<%'######################################
elseif request("type")="admin" then
   
    if request("barname")="" then
       response.write("<script>alert('參數錯誤:網吧名稱不能為空！');history.go(-1)</script>")
    end if%>
<%set rs=server.createobject("adodb.recordset")
 sql="select * from bar where barname='"&request("barname")&"'"
 rs.open sql,conn,1,1
      session("barlogin")=request("barname")
rs.close  
set rs=nothing
response.redirect "index.asp?show="&request("barname")&""
%>
<%'######################################
elseif request("type")="pwd" and sjjh_grade=10 then
    if request.form("pwd")="" or request.form("user")="" then
       response.write("<script>alert('用戶名和密碼都不能為空！');history.go(-1)</script>")
    end if
set rs=server.createobject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,3
rs("user")=request.form("user")
rs("pwd")=request.form("pwd")
rs.Update
rs.close 
set rs=nothing
response.write("<script>alert('密碼修改完成！')</script>")
response.write "<meta http-equiv='Refresh' content='0; URL=index.asp?type=admin'>"
%>
<%'######################################
elseif request("type")="edit" then%>
<!--#include file="char.inc"-->
<%if session("barlogin")=request("barname") then
    if request.form("pwd")="" then
       response.write("<script>alert('密碼不能為空！');history.go(-1)</script>")
    end if
set rs=server.createobject("adodb.recordset")
sql="select * from bar where barname='"&request("barname")&"'"
rs.open sql,conn,1,3
rs("city")=htmlencode(request.form("city"))
rs("typical")=htmlencode(request.form("typical"))
rs("number")=htmlencode(request.form("number"))
rs("add")=htmlencode(request.form("add"))
rs("people")=htmlencode(request.form("people"))
rs("ip")=htmlencode(request.form("ip"))
rs("qdlm")=false
rs("phone")=htmlencode(request.form("phone"))
rs("email")=htmlencode(request.form("email"))
rs("zip")=htmlencode(request.form("zip"))
rs("homepage")=htmlencode(request.form("homepage"))
rs("intros")=htmlencode(request.form("intros"))

rs.Update
rs.close 
set rs=nothing
response.write("<script>alert('你的資料已成功修改！')</script>")
response.write "<meta http-equiv='Refresh' content='0; URL=index.asp?show="&request("barname")&"'>"
else
response.write("<script>alert('登陸超時！')</script>")
response.write "<meta http-equiv='Refresh' content='0; URL=index.asp?type=login&barname="&request("barname")&"'>"
end if%>
<%'######################################
else
response.write("<script>alert('■非法操作！')</script>")
response.write "<meta http-equiv='Refresh' content='0; URL=index.asp'>"
end if%>
<%conn.close
set conn=nothing%>