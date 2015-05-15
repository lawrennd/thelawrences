<%@ LANGUAGE="VBSCRIPT" %>
<%
 Option Explicit
 Dim ranWizard, intID, i, background, theme, frameHeight, pagePicture, pageText, pageType, pagewords

 If myinfo.Theme <> "" Then
	Theme = myinfo.Theme
%>
<!--
	$Date: 9/05/97 11:21a $
	$ModTime: $
	$Revision: 8 $
	$Workfile: default.asp $
-->
	<HTML>
	<HEAD>
	<!--#include virtual ="/iissamples/homepage/sub.inc"-->
	<META NAME="GENERATOR" Content="Microsoft Visual InterDev 1.0">
	<META HTTP-EQUIV="Content-Type" content="text/html; charset=iso-8859-1">
	<TITLE><% call Title %></TITLE>
	<% call styleSheet %>
	<meta name="Microsoft Border" content="tl, default">
</HEAD>
	<BODY TopMargin="0" Leftmargin="0"><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td>

<h1 align="center">
</h1>

<p align="center">
</p>
<p align="center">&nbsp;</p>

</td></tr><!--msnavigation--></table><!--msnavigation--><table border="0" cellpadding="0" cellspacing="0" width="100%"><tr><td valign="top" width="1%">

<p>


</p>

</td><td valign="top" width="24"></td><!--msnavigation--><td valign="top">
	<!--#include virtual ="/iissamples/homepage/theme.inc"-->
	<!--msnavigation--></td></tr><!--msnavigation--></table></BODY>
	</HTML>
<%
 Else
	response.redirect "/IISSamples/Default/welcome.htm"
 End If
%>