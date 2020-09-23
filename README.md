<div align="center">

## Picture browser\(FSO\)


</div>

### Description

The code Desplays a list of images in a selected folder on the clients computer and lets him browse it.

it is not very needed but is a gooood exemple of using fso. one of the thing i tell me studets is to try and make it a txt viewer instead of a image viewer, try it!
 
### More Info
 
in order to run the code u must have a FSO activeX object


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Michael Seltenreich](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-seltenreich.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VbScript \(browser/client side\)

**Category**       |[Files](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files__4-2.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/michael-seltenreich-picture-browser-fso__4-8023/archive/master.zip)





### Source Code

```
<script language="VBScript">
'Sets the Varibles
Dim filesys, demofolder, fil, filecoll, filist, filea,FPath,a
'Creats the FSO object
Set filesys = CreateObject("Scripting.FileSystemObject")
'Pops a box asking for the location of the folder, and keeps it in the varible "a"
a=inputbox("Choose a folder u want to browse.")
FPath=a
'calls the folder the user asked for
Set demofolder = filesys.GetFolder(FPath)
'calls the files in the above folder
Set filecoll = demofolder.Files
'a subroutine that gets a value.
sub ShowPic(path)
'desplays the image accepted in the sub's var in the main screen
viewer.innerHTML="<img alt=" & path & " src='" & a & "\" & path & "'><br>" & path
'ends the sub
end sub
'Creats the basic scenerio for the code
document.write "<table><tr><td><div style=height:600;overflow:scroll><table border=2>"
'runs a loop for each file in the folder
For Each fil in filecoll
'....
filea=filesys.GetExtensionName(fil.name)
	'Checks if the file has an image extension
	if filea="jpg" or filea="gif" or filea="bmp" then
	'if the file is in image he is added to the list of images
 filist = filist & "<tr><td><img alt=" & fil.name & " width=200 height=200 src='" & a & "\" & fil.name & "' onclick=ShowPic('" & fil.name & "')></td></tr>"
 end if
Next
'the full list of images is desplayed
document.write (filist)
'Finals the code and creats the main screen
document.write "</table></div></td><td bgcolor=blue width=100% align=center valign=center><div name=viewer id=viewer></div></td></tr></table>"
</script>
```

