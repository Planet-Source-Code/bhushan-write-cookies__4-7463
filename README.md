<div align="center">

## write cookies


</div>

### Description

When a page is loaded, the browser uses a piece of memory to store the pages information. When the user exists the browser, this information can be written to a flat file on the client machine or on the server. This flat file is called a 'cookie'. Generally cookies are harmless, they do not contain any executable code. By retaining information, the cookie can assist the user to have a better web experience by retaining settings and information from the last time that the user was one the page.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bhushan\-](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bhushan.md)
**Level**          |Beginner
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bhushan-write-cookies__4-7463/archive/master.zip)





### Source Code

```
When a page is loaded, the browser uses a piece of memory to store the pages information. When the user exists the browser, this information can be written to a flat file on the client machine or on the server. This flat file is called a 'cookie'. Generally cookies are harmless, they do not contain any executable code. By retaining information, the cookie can assist the user to have a better web experience by retaining settings and information from the last time that the user was one the page.
Here is a simple example which sets a cookie and stores a value called last visit. try it out
<%@ LANGUAGE="VBScript" %>
<% Response.Buffer = true %>
<html>
<head>
<title>Cookies in ASP</title>
</head>
<body bgcolor="#ffffff">
<center>
<table border=0 cellpadding=8 cellspacing=0 width=550>
<tr><td align=center colspan=2>
<hr noshade>
<font color="#cc6600" face="Arial,Helvetica" size=4><b>Cookies in ASP</b></font>
<hr noshade>
</td></tr>
<tr>
<td bgcolor="#cccccc" width=100>&nbsp;</td>
<td width=450>
<font face="Arial,Helvetica" size=2>
When this script runs, it first checks for a cookie named 'lastvisit'. If
found, it displays a message showing the date and time stored in that cookie.
Then it takes the current date and time and saves them in this cookie.
<p>
If you reload the page or visit it again later, it will display a timestamp,
showing you when you were last here (try hitting your RELOAD button if you
don't see it on the next line).
<p>
<% 'Get date and time of last visit from cookie.
last = Request.Cookies("lastvisit")
if last <> "" then
Response.Write("<b>You last visited this page on " & last & ".</b>")
end if %>
<p>
One important thing to note in the source is the line at the beginning of the
page.
<p>
<pre>
Response.Buffer = true
</pre>
<p>
This tells the server to run all the script code in the page before sending any
data to the browser. By default, script output is not buffered, and the server
may start transmitting the page to the browser even as the code is still being
executed.
<p>
Here we want to set cookie information, which is stored as part of the response
header. Everytime a file is transmitted over the web, a short header is
attached to the front of it. When a web server sends a page to a browser this
header contains, among other things, data which tells the browser what kind of
file it is (image, HTML, plain text, etc.). It can also contain cookie data for
the browser to store.
<p>
Since the code that actually sets the cookie value is farther down in the page,
the output needs to be buffered so the cookie can be set before the response
header is sent. Therefore, it is often a good idea to set buffering on anytime
you use cookies within your scripts.
<p>
Also note these two lines.
<p>
<pre>
last = Request.Cookies("lastvisit")
...
Response.Cookies("lastvisit") = Now
</pre>
<p>
The Request object contains data found in the request header sent by the
browser. This will include any cookie data, form input and browser info. The
Response object represents the data that will from the web server back to the
browser, including the header and page contents.
<p>
<font color="#cc6600"><b>Source</b></font>
<p>
<ul>
<li><a href="cookies.asp?source=cookies.asp">cookies.asp</a>
</ul>
<p>
</font>
</td>
</tr>
<tr><td align=center colspan=2>
<hr noshade size=1>
<font face="Arial,Helvetica" size=2><a href="/" target="_top">Home</a></font>
</td></tr>
</table>
<% 'Save current date and time in cookie
Response.Cookies("lastvisit") = Now
Response.Cookies("lastvisit").Expires = DateAdd("d", 30, Date) %>
</center>
</body>
</html>
```

