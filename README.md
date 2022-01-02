# VB-Help
Help in Excel VB
I need Help!
I am making this program from a source that is online.
 I have got to the point that I am putting the code in.
At this point I should have the code make the dark blue line move when I click a different cell it moves to that cell I Clicked and the number in B2 Changes to that Number cell (EX: It is on D 14, I click D15 it should move the Blue Line to that cell and in B2 Should Show 15 in the Sel Row)

 



VB Code:
Option Explicit
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
If Target.Count > 1 Then Exit Sub
If Not Intersect(Target, Range("D13:J9999")) Is Nothing And Range("D" & Target.Row).Value <> Empty Then Range("B2").Value = Target.Row
'Cont_Load
End If

End Sub
Here is the program and the First Code.
 

Source From
Excel for Freelanceres
Program: Contact manager.xlsm


Had Pictures But I guess it is not Visible here
[I need Help With Excel Coding.docx]
(https://github.com/John-zarraz/VB-Help/files/7799187/I.need.Help.With.Excel.Coding.docx)
