<%@ Language="VBScript" CodePage="65001" %>
<!--#include file="config/dbConnection.asp"-->
<!--#include file="config/security.asp"-->
<%
' ============================================================================
'  logout.asp - Destruye la sesion y regresa al login
' ============================================================================

Session.Contents.RemoveAll()
Session.Abandon()
Response.Redirect GetAppBaseUrl() & "login.asp?msg=logout"
Response.End
%>
