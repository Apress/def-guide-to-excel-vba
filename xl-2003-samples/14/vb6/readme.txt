Chart

ExcelChart.exe demonstrates how to use Excel functions
within a Visual Basic program. If you have never
before executed an VB6 program on your PC you
must use setup\setup.exe to install the program
with all necessary libaries. Then execute
Start|Programms|ExcelChart|ExcelChart.

ClipBoard

ClipBrd.dll demonstrates how to use a new ActiveX component
in Excel. ClipBrd.dll was created using Visual Basic 6.
You must first execute clipboard\Setup\Setup.exe! Then 
use ClipBoard\ActiveX_Clip.xls to test the sample.

-------

This information applies to both samples:

During the setup process, it might be that you have to restart
your system to allow the replacement of certain system DLLs.
This is due to the complicated setup process for VB6 projects.

This examples were tested on Windows NT 4 SP 4 english with 
Office 2000 english (no SPs), and they did work. They were also
tested with Windows 2000 english (no SP) with Office 2000 english (no SP).

Even if the setup process succeeds, it might happen that the
example still does not work as expected, depending on which 
language of Office 2000 and which Service Pack you have installed.
In this cases you would need Visual Basic 6 to recompile the samples.
In case of the clipboard example, you must also delete and then create 
anew the reference to the clipboard.dll in the ActiveX_Clip.dll.

I am sorry for this 'DLL hell'; it is not my fault (but M$'s). 
I am not able to provide further support for these samples. 

Sorry about that,

	Michael Kofler