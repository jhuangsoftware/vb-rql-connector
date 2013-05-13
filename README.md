vb-rql-connector
================

All version rql connector wrriten in VB

Usage
=====
<!-- #include file="VbRqlConnector.asp" -->
<%
Dim oVbRqlConnector

'Instantiate object
Set oVbRqlConnector = New VbRqlConnector

oVbRqlConnector.InitializeConnectionType()
oVbRqlConnector.SendRql("###Your RQL HERE###")

'Destroy object
Set oGreeting = nothing
%>
