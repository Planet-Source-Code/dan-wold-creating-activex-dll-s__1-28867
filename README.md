<div align="center">

## Creating ActiveX DLL's


</div>

### Description

This Articles shows you how to create a Simple ActiveX DLL and reference it to your Project.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Wold](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-wold.md)
**Level**          |Beginner
**User Rating**    |4.9 (39 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-wold-creating-activex-dll-s__1-28867/archive/master.zip)





### Source Code

<HTML>
<HEAD><TITLE>Creating ActiveX DLL's</TITLE></HEAD>
<BODY>
<H1 ALIGN=CENTER>Creating a simple ActiveX DLL</H1>
<BR><BR><BR><BR>
<CENTER>
<P><B>To create a simple ActiveX DLL to use with your program follow these instructions</P></B>
</CENTER>
<PRE>
Step 1: Open Visual Basic, For the New Project, Select "ActiveX Dll"
Step 2: Rename the Class Module "Class1" to "Math" You will be calling this class later.
Step 3: Goto the menu and select, Project > Project Properties
Step 4: Change the Project name to MathFuncDll
Step 5: Change the Project Description to "Simple Math Functions" And click "OK"
Step 6: In the Class Module (Math) Put the following code:
</PRE>
<BR><BR>
<PRE>
Option Explicit
Public Function Add(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
Add = FirstNumber + SecondNumber
End Function
Public Function Subtract(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
Subtract = FirstNumber - SecondNumber
End Function
Public Function Divide(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
Divide = FirstNumber / SecondNumber
End Function
Public Function Multiply(ByVal FirstNumber As Long, ByVal SecondNumber As Long)
Multiply = FirstNumber * SecondNumber
End Function
</PRE>
<BR><BR>
<PRE>
Step 7: Now goto the menu "File > Make MathFuncDll.dll" And Compile your ActiveX dll
</PRE>
<BR><BR>
<H1 ALIGN=CENTER>Congrats If it compiled correctly you have just created your first ActiveX Dll!</H1>
<H1 ALIGN=CENTER>If it didnt compile make sure your code is exactly like mine..</H1>
<BR><BR>
<CENTER>
Question: How do I use this ActiveX DLL now?
Answer: Follow the rest of the steps ;)
</CENTER>
<PRE>
Step 8: Open a New Project, This time select a New "Standard EXE"
Step 9: Now goto menu "Project > References" And click "Browse"
Step 10: Now Browse for your Newly Created DLL And select it. Click "OK"
Step 11: Click "OK" Again to add the referance to your Project.
Step 12: Now in the Form Put the following Code:
</PRE>
<BR><BR>
<PRE>
Option Explicit
'Creates The Object Reference
Dim objNew As MathFuncDll.Math
Private Sub Form_Load()
  'Sets objNew to the new Object referance
  Set objNew = New MathFuncDll.Math
  MsgBox objNew.Add(2, 4)
  MsgBox objNew.Subtract(5, 3)
  MsgBox objNew.Multiply(5, 2)
  MsgBox objNew.Divide(10, 5)
End Sub
</PRE>
<BR><BR>
Step 13: Run your project.
<PRE> OK, Ok, Now your Probably wondering "HOW The Heck Did he referance that?" Well now.. After you
added the DLL Into the Projects Refereances I called upon them by setting them to an Object. I.e:
</PRE>
<BR>
<B>Dim objNew As MathFuncDll.Math</B>
<PRE>
That Refrenced objNew to the "Math" Class inside MathFuncDll.Dll
And
</PRE>
<BR>
<B>Set objNew = New MathFuncDll.Math</B>
<PRE>
Created the Object Referance.
Now I called upon that referance by using objNew
I.e:
</PRE>
<BR>
<B>MsgBox objNew.Add(2, 4)<BR>
MsgBox objNew.Subtract(5, 3)<BR>
MsgBox objNew.Multiply(5, 2)<BR>
MsgBox objNew.Divide(10, 5)<BR>
</B>
<I>objNew.Subtract(FirstNumber,SecondNumber) AKA objNew.Subtract(5, 3)</I>
<H5 ALIGN=CENTER> This is My first tutorial, I know, I understand If you couldnt understand it... Err..
Anyways, Im not the best Tech Writer, Heck Im not a tech writer :P. But if you want an example feel free
to email me at <A HREF="mailto:e_man_dan@hotmail.com">E_MAN_DAN@HOTMAIL.COM</A><H5>
</BODY>
</HTML>

