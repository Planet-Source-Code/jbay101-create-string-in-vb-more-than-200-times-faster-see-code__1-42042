<div align="center">

## Create string in VB more than 200 times faster\! See code\!

<img src="PIC20021231165326445.JPG">
</div>

### Description

This article shows how to use the OLE Automation DLL to create a string without relying on VB to do it. When VB creates a string, it automaticly fills it with data (which takes a great deal of time when dealling with large strings). This way bypasses VB and creates the string itself, without filling it with data. The result? 200 times faster! I have updated the tutorial to include Example 1 re-written using the faster method. The code includes a benchmark and the function described below. Please vote or leave a comment!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-12-31 14:07:10
**By**             |[jbay101](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jbay101.md)
**Level**          |Advanced
**User Rating**    |4.9 (69 globes from 14 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Object Oriented Programming \(OOP\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/object-oriented-programming-oop__1-47.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Create\_str15211912312002\.zip](https://github.com/Planet-Source-Code/jbay101-create-string-in-vb-more-than-200-times-faster-see-code__1-42042/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Visual Basic vs</title>
</head>
<body>
<p><b><font face="Verdana">Fast string in Visual Basic, version 2</font></b></p>
<p><font face="Verdana" size="2"><b>Visual Basic vs. C++</b></font></p>
<p><font face="Verdana" size="2">Visual Basic stores it's strings in a type
referred to in C++ as a BSTR. This type is completely different from the C char
type, as a BSTR doesn't necessarily terminate with a null, and it has a
different header. The C char is stored as an array of bytes, terminating at a
null byte or character 0x0. Unlike C or C++, when you create a string in VB it
is automatically filled with data. </font></p>
<p><font face="Verdana" size="2"><b>The SLOW way&nbsp; - Visual Basic's String
creation<br>
</b>When you dynamically create a string in Visual Basic, there are only two
methods that VB supports. These are:<br>
1. Using the String function<br>
&nbsp;&nbsp;&nbsp; <u>Example:<br>
</u></font><font size="2" face="Courier New">&nbsp;&nbsp;&nbsp;
<font color="#000080">Dim</font> strData <font color="#000080">As String&nbsp;&nbsp;&nbsp;
</font><font color="#008000">'our string variable</font><br>
&nbsp;&nbsp;&nbsp; <font color="#000080">Open</font> &quot;test.bin&quot;
<font color="#000080">For Binary Access Read As</font> #1&nbsp;&nbsp;&nbsp;
<font color="#008000">'open a file</font><br>
&nbsp;&nbsp;&nbsp; strData = String(LOF(1), 0)&nbsp;&nbsp;&nbsp;
<font color="#008000">'create a buffer</font><br>
&nbsp;&nbsp;&nbsp; <font color="#000080">Get</font> #1, , strData&nbsp;&nbsp;&nbsp;
<font color="#008000">'read data into the buffer</font><br>
&nbsp;&nbsp;&nbsp; <font color="#000080">Close</font> #1&nbsp;&nbsp;&nbsp;
<font color="#008000">'close the file</font><br>
</font><font face="Verdana" size="2">&nbsp;&nbsp;&nbsp; The String function
takes two parameters, the length of the string and the character to fill the
string with.</font></p>
<p><font face="Verdana" size="2">2. Using the Space function<br>
&nbsp;&nbsp;&nbsp; This is much like using the String function, except it
automatically fills the string with spaces.</font></p>
<p><font face="Verdana" size="2">Now, for the example above all we want is an
empty storage space to fill with data. But VB doesn't do this. In both
instances, VB fills the string with data, which can take a lot of time. This is
where the API optimization comes into play.</font></p>
<p>&nbsp;</p>
<p><font face="Verdana" size="2"><b>The FAST way - the OLE Automation library<br>
</b>The OLE Automation library provides support, not only for the BSTR type but
also for all variable-related operations. To increase the speed of the string
creation, we want to tell the OLE Automation library to create a region of
memory that we can access - without filling it with data. To do this we will use
two functions, RtlMoveMemory in the windows kernel and SysAllocStringByteLen is
the OLE Automation library. The declarations are below.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Declare Sub</font>
RtlMoveMemory <font color="#000080">Lib</font> &quot;kernel32&quot; (dst<font color="#000080">
As Any</font>, src<font color="#000080"> As Any</font>, <font color="#000080">
ByVal</font> nBytes&amp;)<br>
<font color="#000080">Declare Sub</font> SysAllocStringByteLen&amp;
<font color="#000080">Lib</font> &quot;oleaut32&quot; (<font color="#000080">ByVal</font>
olestr&amp;, <font color="#000080">ByVal</font> BLen&amp;)</font></p>
<p><font face="Verdana" size="2">The RltMoveMemory function copies nBytes bytes
from the src address to the dst address. The SysAllocStringByteLen allocates
BLen of storage space for a BSTR, or in this case a Visual Basic String. In
reality, the Visual Basic String is nothing more than a pointer, or a reference
to an address in memory that can be used to store the data. With this in mind,
we can create out own string allocation function, as shown below.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Public Function
</font>AllocString(ByVal lSize <font color="#000080">As Long</font>)
<font color="#000080">As String</font><br>
RtlMoveMemory <font color="#000080">ByVal</font> <font color="#000080">VarPtr</font>(AllocString_ADVANCED),
SysAllocStringByteLen(0&amp;, lSize + lSize), 4&amp;<br>
<font color="#000080">End Function</font><br>
<br>
</font><font face="Verdana" size="2">This may look a bit complicated at first
but it is really relatively simple. The function allocates the space and then
copies the 4 byte pointer from this space to the string returned by the
function. If we were to expand the function a little it would look like this:</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Public Function
</font>AllocString(ByVal lSize <font color="#000080">As Long</font>)
<font color="#000080">As String</font><br>
<font color="#000080">Dim</font> lPtr <font color="#000080">As Long&nbsp;&nbsp;&nbsp;
</font><font color="#008000">'the address of the allocated memory</font><br>
<font color="#000080">Dim</font> lRetPtr<font color="#000080"> As Long&nbsp;&nbsp;&nbsp;
</font><font color="#008000">'the pointer to the return variable<br>
</font><font color="#000080">Dim</font> sBuffer <font color="#000080">As String&nbsp;&nbsp;&nbsp;
</font><font color="#008000">'the variable to return</font><br>
lRetPtr<font color="#000080"> = VarPtr</font>(sBuffer)&nbsp;&nbsp;&nbsp;
<font color="#008000">'the pointer to the string buffer</font><br>
lPtr = SysAllocStringByteLen(0&amp;, lSize + lSize)&nbsp;&nbsp;&nbsp;
<font color="#008000">'allocate the memory and get it's pointer</font><br>
RtlMoveMemory <font color="#000080">ByVal</font> lRetPtr, lPtr, 4&amp;&nbsp;&nbsp;&nbsp;
'<font color="#008000">copy the pointer address</font><br>
AllocString = sBuffer&nbsp;&nbsp;&nbsp; <font color="#008000">'return the string
with the modified pointer</font><br>
<font color="#000080">End Function</font><br>
<br>
</font><font face="Verdana" size="2">As someone highlighted in the previous
tutorial, when a value is returned it is duplicated and added to the stack. When
the function ends, this value is pushed off the stack and return to the assigned
variable. However, this is where some more knowledge of how VB works is
required. Visual Basic is not duplicating the data. All that Visual Basic is
doing is duplicating the pointer to the data. Why move 30 MB when you can move 4
bytes? Still, returning a value does take time, and if you a looking for a few
more miliseconds you could try making the call inline (removing the function all
together). For example, if your string is called strBuffer you could use the
code below.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Dim</font> strBuffer
<font color="#000080">As String</font></font><br>
<font face="Courier New" size="2">RtlMoveMemory <font color="#000080">ByVal</font> <font color="#000080">VarPtr</font>(strBuffer),
SysAllocStringByteLen(0&amp;, 100 + 100), 4&amp; <font color="#008000">'
allocate 100 bytes</font><br>
</font><br>
<font face="Verdana" size="2">This method will be slightly faster, but I don't
think it's worth the trouble (unless you only need to allocate the data once)</font></p>
<p><font face="Verdana" size="2">As most of you know, when dealing with the
API it is very important to free all the memory you allocate, otherwise you can
easily develop memory leaks. But the best part of using the above method is that
we don't have to worry about freeing the memory block. When your Visual Basic
program ends (of a function/sub containing the relative variable ends), VB
automatically checks each variable and frees the memory associated with them. But
this is not a VB variable you may say? Wrong. This is a normal VB string
variable, we have just created it without VB. Any Visual Basic string function
will still work on the data. VB just doesn't know how it was allocated - but VB
doesn't care. If you really wanted, you could write a small function to delete a
string. Just be careful about how you do it. Since to create the variable we
just copied the pointer, some people may think that the below code would free
the string.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Public Function
</font>DeallocString(sString <font color="#000080">As String</font>)<br>
<font color="#000080">Dim</font> lPtr <font color="#000080">As Long&nbsp;&nbsp;&nbsp;
</font><font color="#008000">'the address of the allocated memory</font><br>
lPtr<font color="#000080"> = VarPtr</font>(sBuffer)&nbsp;&nbsp;&nbsp;
<font color="#008000">'the pointer to the string buffer</font><br>
RtlMoveMemory <font color="#000080">ByVal</font> lPtr, 0&amp;, 4&amp;&nbsp;&nbsp;&nbsp;
'<font color="#008000">copy the pointer address (nulls)</font><br>
<font color="#000080">End Function</font><br>
<br>
</font><font face="Verdana" size="2">When dealing with other API types (and VB
types), erasing the pointer will tell VB or the API that the variable hasn't
been initialized. But VB will loose track of all the memory associated with the
string in this case. For the enclosed sample, that is 30 MB or RAM!!! The
correct way to remove the string is to use VB to do it. The easiest way is to
assign its value to &quot;&quot;. But if you really MUST write a function, you could tri
the one below.</font></p>
<p><font face="Courier New" size="2"><font color="#000080">Public Function
</font>DeallocString(sString<font color="#000080"> As String</font>)<br>
sString = &quot;&quot;<font color="#000080"><br>
End Function</font></font></p>
<p><font face="Verdana" size="2">Sometimes the simplest way is the best!</font><font face="Verdana" size="2" color="#000080">
</font><font face="Courier New" size="2"><br>
<br>
</font><font face="Verdana" size="2">Now that we know how to allocate strings
the fast way, we can re-write the sample in Example 1.</font></p>
<p><font size="2" face="Courier New">&nbsp;&nbsp;&nbsp;
<font color="#000080">Dim</font> strData <font color="#000080">As String&nbsp;&nbsp;&nbsp;
</font><font color="#008000">'our string variable</font><br>
&nbsp;&nbsp;&nbsp; <font color="#000080">Open</font> &quot;test.bin&quot;
<font color="#000080">For Binary Access Read As</font> #1&nbsp;&nbsp;&nbsp;
<font color="#008000">'open a file</font><br>
&nbsp;&nbsp;&nbsp; strData = AllocString(LOF(1))&nbsp;&nbsp;&nbsp;
<font color="#008000">'create a buffer using out function</font><br>
&nbsp;&nbsp;&nbsp; <font color="#000080">Get</font> #1, , strData&nbsp;&nbsp;&nbsp;
<font color="#008000">'read data into the buffer</font><br>
&nbsp;&nbsp;&nbsp; <font color="#000080">Close</font> #1&nbsp;&nbsp;&nbsp;
<font color="#008000">'close the file</font><br>
</font><font face="Verdana" size="2"><br>
It really is not that difficult, and it makes a HUGE speed increase. This
article comes with the above function and a benchmark to show the dramatic speed
difference.</font></p>
<p><font face="Verdana" size="2">The next tutorial will talk about making string functions (compare,
join etc) as fast as C and will show how to make a C string in Visual Basic.
Please leave a comment or vote!</font></p>
</body>
</html>

