VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
    Dim name As String 'the current name row being checked
    Dim Students As Integer 'the total number of students
    Dim test As Boolean
    Dim count As Integer
    Dim GradeCard As Integer
    Dim delta As Integer
    Dim last As Integer
    Dim current As Integer
    Dim countminus As Integer
    

    count = 4
    Sheets("MP Data").Activate
    Students = Sheet21.Cells(1, 3) - 2
        
    Sheets("Grade Cards").Activate
    ActiveSheet.ResetAllPageBreaks
    GradeCard = 1
    delta = 0
    test = False
    
    'unmerge cells
    arange = 1
    brange = Students * 25
    trange = "A" & arange & ":H" & brange


ActiveSheet.Range(trange).UnMerge
    Do
        countminus = count - 1
        last = Application.WorksheetFunction.CountA(Sheet11.Range(("C3:C" & countminus)))
        current = Application.WorksheetFunction.CountA(Sheet11.Range(("C3:C" & count)))
        
        
        name = Sheet11.Cells(count, 3)
        delta = current - last
        
        
        
        If delta >= 1 Then
            Call gradesheet(count, GradeCard)
    
    
            'Sheets("Gradebook-Roster").Cells(Count, 3).Copy Destination:=Sheets("Test").Cells(GradeCard, 1)
            GradeCard = GradeCard + 1
        Else
        'literally do nothing
        End If
        
        
        count = count + 1
       
        
        
        If GradeCard = Students + 1 Then
            test = True
        ElseIf count = 1500 Then
            test = True
        Else
            test = False
        End If


    Loop Until test = True
    
    
    





End Sub

Private Sub CommandButton10_Click()
Dim name As String
Dim extrahours As Integer
Dim destination As String

name = Sheet1.Cells(10, 5)

extrahours = Sheet1.Cells(7, 5)
destination = Sheet1.Cells(8, 5)

Call addextrahours(name, extrahours, destination)





End Sub
Public Sub addextrahours(name As String, extrahours As Integer, Optional destination As String = "In class", Optional reason As String = "Extra time spent in class on")
Dim slot As Integer
Dim i As Integer
Dim hourcol As Integer
hourcol = 5
extrahours = extrahours - 1


If destination = "DSU" Then hourcol = 6
i = 1


Do
slot = findstudentrow(name, i, True)


rabbits = True

If slot = 201 Then rabbits = False
If i = 9 Then rabbits = True
i = i + 1
Loop Until rabbits = True



If slot = 201 Then
MsgBox ("Could not find student")
Else
Sheet14.Cells(slot, hourcol) = Sheet14.Cells(slot, hourcol) + extrahours
Dim noteentry As Integer
noteentry = Sheet14.Cells(slot, 7).Value + 8
Sheet14.Cells(slot, noteentry) = reason & " (on " & Date & ")"

End If


End Sub

Private Sub CommandButton11_Click()

Call resetweek

Call advancehomework



End Sub

Public Sub resetweek()
'++++++++++++++++++++++++++++++++++++++
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'++++++++++++++++++++++++++++++++++++++

For i = 4 To 100
Sheet15.Cells(i, 10) = Sheet15.Cells(i, 8).Value
Sheet15.Cells(i, 9) = Sheet15.Cells(i, 7).Value
Sheet15.Cells(i, 8) = Sheet15.Cells(i, 6).Value
Sheet15.Cells(i, 7) = Sheet15.Cells(i, 5).Value
Sheet15.Cells(i, 6) = Sheet15.Cells(i, 4).Value
Sheet15.Cells(i, 5) = Sheet15.Cells(i, 3).Value

Next i

End Sub
Public Sub advancehomework()
'++++++++++++++++++++++++++++++++++++++
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'++++++++++++++++++++++++++++++++++++++
Dim Decision As VbMsgBoxResult

Decision = MsgBox("Do you want to advance the homework to the next week?", vbYesNo, "DECIDE")
Select Case Decision
Case 6 'Yes case
For i = 4 To 150
'Call setassignments(i)

Sheet11.Cells(i, 18) = Sheet15.Cells(i, 47).Value




Next i

Case 7 'No case
'Do nothing
End Select



End Sub
Public Sub setassignments(i As Integer)




End Sub

Public Sub CreateGradeReport(name As String, credittype As String, grade As String, subject As String, Optional noteaddendum As String = "", Optional fileaddendum As String = "")
'these are to work the file saving
Dim sourcepath As String
Dim path As String
Dim objWord
Dim mathsubject As String
Dim currentdate As String







'opens the copy of the document




'####
'@@@@
'####


'Inputs the information needed to create the grade report.
'this sets the current student
Dim loopnumber As Integer
Dim studentrow As Integer
Dim loopfinish As Boolean



loopnumber = 1
Do

studentrow = findstudentrow(name, loopnumber, True)


loopnumber = loopnumber + 1

loopfinish = True
If studentrow = 201 Then loopfinish = False
If loopnumber > 8 Then loopfinish = True



Loop Until loopfinish = True

If studentrow = 201 Then

MsgBox "Could not find student! Please check to make sure student has not been removed."
Else
'everthing below here is in this if

Dim newdoc
sourcepath = fRelativePath()
sourcepath = sourcepath & "\Reporthacked.doc"
Set objWord = CreateObject("Word.Application")
Set newdoc = objWord.Documents.Open(sourcepath)
'saves the document as a copy - named correctly
currentdate = Date

currentdate = Replace(currentdate, "/", "-")

path = fRelativePath()
path = path & "\Grade Reports\GradeReport-" & name & "-" & currentdate & fileaddendum & ".doc"


'date
Dim dateRange
Set dateRange = newdoc.Bookmarks("Text31").Range
dateRange.text = currentdate
newdoc.Bookmarks.Add "Text31", dateRange


newdoc.Save

'studentname

Dim nameRange
Set nameRange = newdoc.Bookmarks("Text32").Range
nameRange.text = name
newdoc.Bookmarks.Add "Text32", nameRange


newdoc.Save

'final grade


If grade = "Auto" Then
Dim lgrade As String
Dim sheetgrade As Variant

sheetgrade = Sheet11.Cells(studentrow, 14)

If sheetgrade > 0.93 Then
lgrade = "A"
ElseIf sheetgrade > 0.9 Then
lgrade = "A-"
ElseIf sheetgrade > 0.87 Then
lgrade = "B+"
ElseIf sheetgrade > 0.83 Then
lgrade = "B"
ElseIf sheetgrade > 0.8 Then
lgrade = "B-"
ElseIf sheetgrade > 0.77 Then
lgrade = "C+"
ElseIf sheetgrade > 0.73 Then
lgrade = "C"
ElseIf sheetgrade > 0.7 Then
lgrade = "C-"
ElseIf sheetgrade > 0.67 Then
lgrade = "D+"
ElseIf sheetgrade > 0.63 Then
lgrade = "D"
ElseIf sheetgrade > 0.6 Then
lgrade = "D-"
Else
lgrade = "F"
End If

grade = lgrade

Else
'do nothing the grade = grade
End If

Dim gradeRange
Set gradeRange = newdoc.Bookmarks("Text34").Range
gradeRange.text = grade
newdoc.Bookmarks.Add "Text34", gradeRange



newdoc.Save

'class hours
Dim classhours As Integer
Dim subtractedhours As Integer
Dim currenthours As Integer

currenthours = Sheet11.Cells(studentrow, 19).Value

If credittype = "90 hours" Then
classhours = 90
subtractedhours = 90 * -1
ElseIf credittype = "Early credit" Then
classhours = 90
'this confirms the hours to subtract, edit this



Decision = MsgBox("Do you want to reset student hours to 0?", vbYesNo, "DECIDE")
Select Case Decision
Case 6 'Yes case
subtractedhours = currenthours * -1
Case 7 'No case
Dim myval As Integer
myValue = InputBox("Please enter the number of hours to subtract" & vbCrLf _
& " Current Hours = " & currenthours, "Hours to Subtract", currenthours)





subtractedhours = myValue * -1
End Select






ElseIf credittype = "Current hours" Then
classhours = currenthours
subtractedhours = currenthours * -1
End If

Dim hoursRange
Set hoursRange = newdoc.Bookmarks("Text35").Range
hoursRange.text = classhours
newdoc.Bookmarks.Add "Text35", hoursRange

Dim reason As String
reason = name & " had " & classhours & " added to his transcript on " & Date & ". Therefore " & (subtractedhours * -1) & " were subtracted from his total."
Call addextrahours(name, subtractedhours, "In class", reason)

Dim subtractedcredit As Variant
subtractedcredit = subtractedhours / (-12)
Dim creditslot As Integer
Dim count As Integer
count = 1

For i = 1 To 8
creditslot = findstudentrow(name, count, True)
count = count + 1

If creditslot = 201 Then
'do nothing
Else
Sheet15.Cells(creditslot, 13) = Sheet15.Cells(creditslot, 13) - subtractedcredit
End If

Next i



newdoc.Save

'mathsubject

If subject = "Auto" Then
mathsubject = Sheet11.Cells(studentrow, 15)

    If mathsubject = "Alg" Then
    mathsubject = "Algebra"
    ElseIf mathsubject = "Geo" Then
    mathsubject = "Geometry"
    ElseIf mathsubject = "Alg2" Then
    mathsubject = "Algebra 2"
    ElseIf mathsubject = "HiSET" Then
    mathsubject = "Math 5"
    Else
    mathsubject = "Adv. Math"
    End If


Else
mathsubject = subject
End If




Dim objRange
Set objRange = newdoc.Bookmarks("Text33").Range
objRange.text = mathsubject
newdoc.Bookmarks.Add "Text33", objRange


newdoc.Save



'notes
Dim gradereportnote As String
Dim earlycreditcomment As String

earlycreditcomment = ""

If credittype = "Early credit" Then
earlycreditcomment = name & " has completed all the assignments necessary for the course and has shown mastery of the subject sufficent to earn early credit. "
End If


gradereportnote = name & " has spent " & currenthours & " hours in class working on " & mathsubject & ". " & earlycreditcomment & "During this time " & name & " earned a cumulative grade of " & grade & " for " & mathsubject


Dim noteRange
Set noteRange = newdoc.Bookmarks("Text39").Range
noteRange.text = gradereportnote
newdoc.Bookmarks.Add "Text39", noteRange



newdoc.Save


'saves the document


'saves the information to the copy
newdoc.SaveAs2 (path)




'saves a copy of the document for sending and for records

path = fRelativePath()
path = path & "\To Email\GradeReport-" & name & " - " & currentdate & ".doc"
newdoc.SaveAs2 (path)
newdoc.Close

'this whole part was in an if
End If
Decision = MsgBox("Do you want to reset student grades?", vbYesNo, "DECIDE")
Select Case Decision
Case 6 'Yes case
Call recordandresetstudent(name)
Case 7 'No case
'do nothing
End Select




End Sub





Private Sub CommandButton12_Click()

Dim name As String
Dim credittype As String
Dim grade As String
Dim subject As String

name = Sheet1.Cells(23, 5)
credittype = Sheet1.Cells(25, 5)
grade = Sheet1.Cells(28, 5)
subject = Sheet1.Cells(30, 5)


Call CreateGradeReport(name, credittype, grade, subject)



End Sub

Private Sub CommandButton2_Click()

Dim student As String
student =

Call recordandresetstudent(student)



'Dim Decision As VbMsgBoxResult

'Decision = MsgBox("Do you want to advance the homework to the next week?", vbYesNo, "DECIDE")
'Select Case Decision
'Case 6 'Yes case
'MsgBox ("hello world")

'Case 7 'No case
'Do nothing
'End Select

'Dim test As Boolean

'Call placeintobackup(5, 1)

'Call placestudentinrow(150)


'Call clearrow(150)



'test = IsEmpty(Sheet20.Cells(1, 1))
'MsgBox test


'test = IsEmpty(Sheet20.Cells(1, 2))
'MsgBox test


'Dim text As String

'text = Sheet20.Cells(1, 1)

'If text = "False" Then
'MsgBox text
'Else
'do nothing
'End If



'Dim test As Integer
'Dim name As String




'test = findstudentrow(name, 1)

'MsgBox test
'name = "spongee"

'test = findstudentrow(name, 1)

'MsgBox test
 
'name = "wtf?"

'name = Sheet1.Cells(23, 2)

'MsgBox name

'name = Sheet1.Cells(23, 2).Value
 
'MsgBox name

 'name = "kid"
'test = findstudentrow(name, 7, True)

'MsgBox test


'Call placeintobackup(4, 1)

'Call placeintolongtermstorage(1, "Algebra 2 first credit")




'Dim student As String

'Dim path As String

'path = fRelativePath()

'student = "Tom Tester"

'MsgBox student


'MsgBox path

'path = path & "\" & student & ".txt"



'MsgBox path

'Set fs = CreateObject("Scripting.FileSystemObject")
 '   Set a = fs.CreateTextFile(path, True)
  '  a.WriteLine (student)
   ' a.Close

'Call placeintobackup(4, 1)







'Dim Students As Integer
'Dim lgrade As Variant
'Dim lgradecell As Variant
'Dim studentrow As Integer
'Dim gradepercent As Variant
'Dim gradestring As String

'Dim scount As Integer

'scount = Application.WorksheetFunction.CountA(Sheet12.Range("B:B"))

'MsgBox scount
 
'the following is a the testing of the grade card prinitng function

'Sheets("Grade Cards").Activate

'Students = Application.CountA(Cells(4, 2), Cells(1000, 2))

'Sheets("Test").Activate

'Sheet10.Cells(2, 1) = Students

'Students = Sheet21.Cells(1, 3)

'Sheet10.Cells(3, 1) = Students
'Call gradesheet(6, 2)

'the following is the testing of the hours function.


'MsgBox "This adds test"


'studentrow = 4
'lgrade = Worksheets("Gradebook-Roster").Cells(studentrow, 14).Value
'Sheet3.Cells((gradesheetindex + 5), 6)



'If lgrade = "Pass" Then
    'lgradecell = 13
'Else
    'If lgrade >= 0.93 Then
        'lgradecell = 1
    'ElseIf lgrade >= 0.9 Then
        'lgradecell = 2
    'ElseIf lgrade >= 0.87 Then
        'lgradecell = 3
   'ElseIf lgrade >= 0.83 Then
        'lgradecell = 4
    'ElseIf lgrade >= 0.8 Then
        'lgradecell = 5
   ' ElseIf lgrade >= 0.77 Then
       ' lgradecell = 6
  '  ElseIf lgrade >= 0.73 Then
       ' lgradecell = 7
  '  ElseIf lgrade >= 0.7 Then
       ' lgradecell = 8
  '  ElseIf lgrade >= 0.67 Then
       ' lgradecell = 9
   ' ElseIf lgrade >= 0.63 Then
       'lgradecell = 10
   ' ElseIf lgrade >= 0.6 Then
       ' lgradecell = 11
   ' Else
        'lgradecell = 12
    'End If
'End If

'gradepercent = (Worksheets("Gradebook-Roster").Cells(studentrow, 14) * 100)
'gradestring = gradepercent
'gradestring = Left(gradepercent, 5)

'Sheet10.Cells(2, 2) = Sheet10.Cells(2, 2) & " [" & Worksheets("Gradebook-Roster").Cells(studentrow, 3) & "'s Cumulative grade is " & gradestring & "%" & " which is a " & Worksheets("MP Data").Cells(lgradecell, 2) & ". " & Worksheets("Gradebook-Roster").Cells(studentrow, 3) & " currently has " & Worksheets("Gradebook-Roster").Cells(studentrow, 19) & " hours earned toward a math credit." & "]"


End Sub



Public Sub gradesheet(studentrow As Integer, sheetnumber As Integer)

Dim gscolumn As Integer
Dim gsrow As Integer
Dim gradesheetindex As Integer
Dim gradesheetindexend As Integer
Dim Rows As Integer
Dim templaterow As Integer
Dim dummy As Integer
Dim newpage As Integer
Dim lgrade As Variant
Dim lgradecell As Integer
Dim upgrades As Integer
Dim Ccounts As Integer
Dim fl1 As Integer
Dim updown As Integer
Dim arange As Integer
Dim brange As Integer
Dim trange As String
Dim gradepercent As Variant
Dim gradestring As Variant


templaterow = 1

gradesheetindex = (1 + (24 * sheetnumber) - 24)
gradesheetindexend = (gradsheetindex + 24)
newpage = gradesheetindexend + 1
rows2 = gradesheetindex

For Rows = 1 To 24
       
    For gscolumn = 1 To 8
        
        Sheets("Grade Template").Cells(templaterow, gscolumn).Copy destination:=Sheets("Grade Cards").Cells(rows2, gscolumn)
    Next gscolumn
        
    templaterow = templaterow + 1
    rows2 = rows2 + 1
Next Rows



'place student name
'Sheets("Gradebook-Roster").Cells(Studentrow, 2).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 3), 2)

Worksheets("Gradebook-Roster").Cells(studentrow, 3).Copy
Worksheets("Grade Cards").Cells((gradesheetindex + 3), 2).PasteSpecial Paste:=xlPasteValues
'place student class period
'Sheets("Gradebook-Roster").Cells(Studentrow, 1).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 3), 7)

Worksheets("Gradebook-Roster").Cells(studentrow, 1).Copy
Worksheets("Grade Cards").Cells((gradesheetindex + 3), 7).PasteSpecial Paste:=xlPasteValues
'Place week ending
'Sheets("Scripts").Cells(3, 5).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 4), 2)

Worksheets("Scripts").Cells(3, 2).Copy
Worksheets("Grade Cards").Cells((gradesheetindex + 4), 2).PasteSpecial Paste:=xlPasteValues
'Place cottage
'Sheets("Gradebook-Roster").Cells(Studentrow, 4).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 4), 6)

Worksheets("Gradebook-Roster").Cells(studentrow, 5).Copy
Worksheets("Grade Cards").Cells((gradesheetindex + 4), 6).PasteSpecial Paste:=xlPasteValues
'place grade percent
'Sheets("Gradebook-Roster").Cells(Studentrow, 13).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 5), 6)

Worksheets("Grade Calculations").Cells(studentrow, 11).Copy

Worksheets("Grade Cards").Cells((gradesheetindex + 5), 6).PasteSpecial Paste:=xlPasteValuesAndNumberFormats
'place letter grade
lgrade = Worksheets("Grade Calculations").Cells(studentrow, 11).Value
'Sheet3.Cells((gradesheetindex + 5), 6)

If lgrade = "Pass" Then
    lgradecell = 13
Else
    If lgrade >= 0.93 Then
        lgradecell = 1
    ElseIf lgrade >= 0.9 Then
        lgradecell = 2
    ElseIf lgrade >= 0.87 Then
        lgradecell = 3
    ElseIf lgrade >= 0.83 Then
        lgradecell = 4
    ElseIf lgrade >= 0.8 Then
        lgradecell = 5
    ElseIf lgrade >= 0.77 Then
        lgradecell = 6
    ElseIf lgrade >= 0.73 Then
        lgradecell = 7
    ElseIf lgrade >= 0.7 Then
        lgradecell = 8
    ElseIf lgrade >= 0.67 Then
        lgradecell = 9
    ElseIf lgrade >= 0.63 Then
        lgradecell = 10
    ElseIf lgrade >= 0.6 Then
        lgradecell = 11
    Else
        lgradecell = 12
    End If
End If





'Sheet11.Cells(13, 2).Value = lgrade
'Sheets("MP Data").Cells(lgradecell, 2).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 5), 7)

Worksheets("MP Data").Cells(lgradecell, 2).Copy
Worksheets("Grade Cards").Cells((gradesheetindex + 5), 7).PasteSpecial Paste:=xlPasteValues

'place comment
'Sheets("Gradebook-Roster").Cells(Studentrow, 12).Copy Destination:=Sheets("Grade Cards").Cells((gradesheetindex + 19), 2)

Worksheets("Gradebook-Roster").Cells(studentrow, 13).Copy
Worksheets("Grade Cards").Cells((gradesheetindex + 19), 2).PasteSpecial Paste:=xlPasteValues



'add additional information to comment


lgrade = Worksheets("Gradebook-Roster").Cells(studentrow, 14).Value
'Sheet3.Cells((gradesheetindex + 5), 6)

If lgrade = "Pass" Then
    lgradecell = 13
Else
    If lgrade >= 0.93 Then
        lgradecell = 1
    ElseIf lgrade >= 0.9 Then
        lgradecell = 2
    ElseIf lgrade >= 0.87 Then
        lgradecell = 3
    ElseIf lgrade >= 0.83 Then
        lgradecell = 4
    ElseIf lgrade >= 0.8 Then
        lgradecell = 5
    ElseIf lgrade >= 0.77 Then
        lgradecell = 6
    ElseIf lgrade >= 0.73 Then
        lgradecell = 7
    ElseIf lgrade >= 0.7 Then
        lgradecell = 8
    ElseIf lgrade >= 0.67 Then
        lgradecell = 9
    ElseIf lgrade >= 0.63 Then
        lgradecell = 10
    ElseIf lgrade >= 0.6 Then
        lgradecell = 11
    Else
        lgradecell = 12
    End If
End If


If lgradecell = 13 Then
gradestring = "Pass"
Else
gradestring = (Worksheets("Gradebook-Roster").Cells(studentrow, 14) * 100)
gradestring = Left(gradestring, 5)
End If

Worksheets("Grade Cards").Cells((gradesheetindex + 19), 2) = Worksheets("Grade Cards").Cells((gradesheetindex + 19), 2) & " [" & Worksheets("Gradebook-Roster").Cells(studentrow, 3) & "'s Cumulative grade is " & gradestring & "%" & " - " & Worksheets("MP Data").Cells(lgradecell, 2) & ". " & Worksheets("Gradebook-Roster").Cells(studentrow, 3) & " currently has " & Worksheets("Gradebook-Roster").Cells(studentrow, 19) & " hours earned toward a math credit." & "]"



'format comment

Sheets("Grade Cards").Activate

arange = gradesheetindex + 19
brange = gradesheetindex + 23
trange = "B" & arange & ":H" & brange


ActiveSheet.Range(trange).Merge


'print upgrades (upgrades)

For fl1 = 1 To 6
Ccounts = 6 + fl1
updown = Worksheets("Gradebook-Roster").Cells(studentrow, Ccounts).Value + 2

Sheets("Grade Template").Cells(32, updown).Copy destination:=Sheets("Grade Cards").Cells((gradesheetindex + 5 + 2 * fl1), updown)

Next fl1



If sheetnumber Mod 2 = 0 Then
Worksheets("Grade Cards").Rows(sheetnumber * 24 + 1).PageBreak = xlPageBreakManual
Else
'do nothing
End If



End Sub


Public Sub resetday()

Sheets("Gradebook-Roster").Activate

For Each chk In ActiveSheet.CheckBoxes
 
  chk.Value = False
Next chk

End Sub

Public Sub recordday()

Dim count As Integer
Dim delta As Integer
Dim Students As Integer
Dim test As Boolean
Dim seathour As Boolean
Dim currenthours As Integer
Dim Record As String
Dim grade As Variant
Dim wgrade As Variant
Dim studentscompleted As Integer
Dim last As Integer
Dim current As Integer
Dim countminus As Integer

studentscompleted = 0


Students = Sheet21.Cells(1, 3)
count = 4
  
  
Do
        countminus = count - 1
        last = Application.WorksheetFunction.CountA(Sheet11.Range(("C3:C" & countminus)))
        current = Application.WorksheetFunction.CountA(Sheet11.Range(("C3:C" & count)))
        
        
      
        delta = current - last
        
        If delta >= 1 Then
 
       seathour = Sheet14.Cells(count, 1)
        
            If seathour = True Then
            
            'this line runs if ther is a student and I need to add the hours.
                Call recordstudent(count)
            'This ends the sequence to add the hours to a students record
                studentscompleted = studentscompleted + 1
                
                 Else
            'quite literally do nothing
             
   
                End If
            Else
            'Do nothing again, because coding is wierd
            End If
            
        count = count + 1
       
        If studentscompleted = Students Then
            test = True
        ElseIf count = 1500 Then
        'This is only so there is no stack overflow, it sets a limit of 1500 students in the class/gradebook
        
            test = True
        Else
            test = False
        End If

    Loop Until test = True







End Sub

Public Sub recordstudent(student As Integer)
Dim currenthours As Integer
Dim count As Integer
Dim grade As Variant
Dim wgrade As Variant
Dim Record As String
Dim Ccounts As Integer
Dim downgrades As Integer
Dim updown As Integer

count = student

currenthours = Sheet14.Cells(count, 7).Value
currenthours = currenthours + 8
grade = Sheet15.Cells(count, 11).Value
wgrade = Sheet15.Cells(count, 2).Value
    
    Record = Now
    'add the date to the  attendance
    Sheet14.Cells(count, currenthours) = Record
    
    'add the current grade, weekly grade, and current downgrades to the same place in subsequent sheets
    Sheet16.Cells(count, currenthours) = Worksheets("Gradebook-Roster").Cells(count, 14).Value
    Sheet17.Cells(count, currenthours) = Worksheets("Grade Calculations").Cells(count, 11).Value
    
  downgrades = 0
    
    For fl2 = 1 To 6
Ccounts = 6 + fl2
updown = Worksheets("Gradebook-Roster").Cells(student, Ccounts).Value

If updown < 3 Then
downgrades = downgrades + (3 - updown)
Else
'do nothing
End If

Next fl2
    
    
Sheet18.Cells(count, currenthours) = downgrades
    
    
    '"Student present on " & Now & " With a current grade of " & grade & " and " & wgrade & " for the week"
    



End Sub

Private Sub CommandButton3_Click()

Call recordday

Call resetday

Call LinkCheckBoxes


End Sub
Public Sub LinkCheckBoxes()
Dim chk As CheckBox
Dim lCol As Long
Dim count As Integer


lCol = 1 'number of columns to the right for link
count = 3

Sheets("Gradebook-Roster").Activate

For Each chk In ActiveSheet.CheckBoxes
 count = count + 1
   With chk
      .LinkedCell = _
      "Attendance!A" & count
   End With



Next chk

End Sub

Public Sub advanceweek()

Do










If count = 1500 Then
test = True
Else
' do nothing
End If


Loop Until test = True



End Sub
Public Sub addstudent(period As Integer, name As String, isold As Boolean)






End Sub

Public Sub removestudent(name As String, note As String)
Dim slots As Integer
Dim period As Integer
Dim rowwithstudent As Integer





'remove period 1

'remove period 2

'remove period 3

'remove period 4

'remove period 5

'remove period 6

'remove period 7

'remove period 8

slots = 0
For i = 1 To 8
period = i
rowwithstudent = findstudentrow(name, period, True)
If rowwithstudent = 201 Then
'do nothing
Else
slots = slots + 1
Call placeintobackup(rowwithstudent, slots)
Call clearrow(rowwithstudent)
End If



Next i

If slots > 0 Then
Call placeintolongtermstorage(slots, note)
Else
'do nothing
End If



End Sub


Public Sub backupstudent(studentrow As Integer)
'this is something else
Dim newrow As Integer
Dim scount As Integer

scount = Application.WorksheetFunction.CountA(Sheet12.Range("B:B"))

newrow = scount




End Sub
'This is the actual start of the student movement functions################
'#######################




Public Sub placeintobackup(studentrow As Integer, backupnumber As Integer)

Dim slot As Integer
Dim col As Integer



slot = 6 * (backupnumber) + 7

'pull from grade - roster


col = 1



Do
'Worksheets("Gradebook-Roster").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues


Sheet20.Cells(slot, col) = Sheet11.Cells(studentrow, col).Value
col = col + 1
Loop Until col = 150

slot = slot + 1
'pull from attendance
col = 1
Do
'Worksheets("Attendance").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues


Sheet20.Cells(slot, col) = Sheet14.Cells(studentrow, col).Value
col = col + 1
Loop Until col = 150
slot = slot + 1

'pull from grade calculations
col = 1
Do
Sheet20.Cells(slot, col) = Sheet15.Cells(studentrow, col).Value
col = col + 1
Loop Until col = 150
slot = slot + 1

'pull from attendance - Cumulative grade
col = 1
Do

'Worksheets("Attendance-C.Grade").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues

Sheet20.Cells(slot, col) = Sheet16.Cells(studentrow, col).Value
col = col + 1
Loop Until col = 150
slot = slot + 1
'pull from attedance - Weekly grade
col = 1
Do

'Worksheets("Attendance-W.Grade").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues

Sheet20.Cells(slot, col) = Sheet17.Cells(studentrow, col).Value
col = col + 1
Loop Until col = 150
slot = slot + 1
'pull from attendance - Downgrades
col = 1
Do

'Worksheets("Attendance-downgrades").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues

Sheet20.Cells(slot, col) = Sheet18.Cells(studentrow, col).Value
col = col + 1
Loop Until col = 150
slot = slot + 1




End Sub
Public Sub calibrateerasecell(studentrow As Integer)
''this does not work yet
'#############################################################


Dim slot As Integer
Dim col As Integer
Dim teststring As String




slot = 1

'pull from grade - roster


col = 1

teststring = Sheet11.Cells(studentrow, 14)

MsgBox teststring


Do
'Worksheets("Gradebook-Roster").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues
    teststring = Sheet11.Cells(studentrow, col)
    teststring = Mid(teststring, 1, 1)
   
        If teststring = "=" Then
        Sheet20.Cells(slot, col) = "False"
        Else
        Sheet20.Cells(slot, col) = ""
        
        End If
   
col = col + 1
Loop Until col = 150

slot = slot + 1
'pull from attendance
col = 1
Do
'Worksheets("Attendance").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues
  teststring = Sheet14.Cells(studentrow, col)
    teststring = Mid(teststring, 1, 1)
    If Sheet20.Cells(slot, col) = False Then
        'Do nothing
    Else
        If teststring = "=" Then
        Sheet20.Cells(slot, col) = "False"
        Else
        Sheet20.Cells(slot, col) = ""
        
        End If
    End If
col = col + 1
Loop Until col = 150

'pull from attendance - Cumulative grade
col = 1
Do

'Worksheets("Attendance-C.Grade").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues

  teststring = Sheet16.Cells(studentrow, col)
    teststring = Mid(teststring, 1, 1)
    If Sheet20.Cells(slot, col) = False Then
        'Do nothing
    Else
        If teststring = "=" Then
        Sheet20.Cells(slot, col) = "False"
        Else
        Sheet20.Cells(slot, col) = ""
        
        End If
    End If
col = col + 1
Loop Until col = 150

slot = slot + 1
'pull from attedance - Weekly grade
col = 1
Do

'Worksheets("Attendance-W.Grade").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues

  teststring = Sheet17.Cells(studentrow, col)
    teststring = Mid(teststring, 1, 1)
    If Sheet20.Cells(slot, col) = False Then
        'Do nothing
    Else
        If teststring = "=" Then
        Sheet20.Cells(slot, col) = "False"
        Else
        Sheet20.Cells(slot, col) = ""
        
        End If
    End If
col = col + 1
Loop Until col = 150

slot = slot + 1
'pull from attendance - Downgrades
col = 1
Do

'Worksheets("Attendance-downgrades").Cells(studentrow, col).Copy
'Worksheets("student-buffer").Cells(slot, col).PasteSpecial Paste:=xlPasteValues

  teststring = Sheet18.Cells(studentrow, col)
    teststring = Mid(teststring, 1, 1)
    If Sheet20.Cells(slot, col) = False Then
        'Do nothing
    Else
        If teststring = "=" Then
        Sheet20.Cells(slot, col) = "False"
        Else
        Sheet20.Cells(slot, col) = ""
        
        End If
    End If
col = col + 1
Loop Until col = 150

slot = slot + 1




End Sub

Public Sub placeintolongtermstorage(numberofentries As Integer, note As String, Optional fileaddendum As String = "Normal")
Dim student As String
Dim entrynumber As Integer
Dim path As String
Dim nextelement As String
Dim textline As String
Dim filedate As String
Dim testforline As Boolean


filedate = Date
filedate = Replace(filedate, "/", "-")
path = fRelativePath()

student = Sheet20.Cells(13, 3)
 

MsgBox student & " will be placed in long term storage"




path = path & "\LTS\" & filedate & " " & student & " " & fileaddendum & ".txt"






note = student & ", " & note

Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CreateTextFile(path, True)
    a.WriteLine (note)
    a.Close

Open path For Output As #1

entrynumber = 1

Do

For i = 1 To 6
rowi = i + entrynumber * 6 + 6

'second for loop

For columni = 1 To 150

testforline = IsEmpty(Sheet20.Cells(rowi, columni))
If testforline = False Then
If IsError(Sheet20.Cells(rowi, columni)) = True Then
nextelement = ""
Else

nextelement = Sheet20.Cells(rowi, columni)
textline = nextelement
End If

End If

Write #1, textline
textline = " "

Next columni



Next i

entrynumber = entrynumber + 1
numberofentries = numberofentries - 1
Loop Until numberofentries = 0


Close #1



End Sub

Function fRelativePath() As String
Dim strCurName As String


strCurName = Application.ActiveWorkbook.path



fRelativePath = strCurName
End Function


Public Sub newstudenttobackup()
Dim newi As Integer
Dim newj As Integer
Dim i As Integer
Dim J As Integer


For i = 1 To 6
newi = 6 + i
bi = 12 + i
For J = 1 To 150
Sheet20.Cells(bi, J) = Sheet20.Cells(newi, J)

Next J
Next i





End Sub

Function findslot(period As Integer)
Dim studentrow As Integer
Dim rowtocheck As Integer
Dim i As Integer
Dim test As Boolean


rowtocheck = 10 * (period - 1) + 4
i = 0
test = False

Do
i = i + 1


test = IsEmpty(Sheet11.Cells(rowtocheck, 3))
If test = True Then
studentrow = rowtocheck
Else
'do nothing
End If


If i = 10 Then
test = True
Else
'do nothing
End If
rowtocheck = rowtocheck + 1
Loop Until test = True






findslot = studentrow

End Function

Function findstudentrow(name As String, period As Integer, Optional suppresserror As Boolean = False)
Dim studentrow As Integer
Dim testrow As Integer
Dim i As Integer
Dim testname As String
Dim errorcatch As Boolean
Dim errormsg As String

testrow = (period - 1) * 10 + 4
Do
i = i + 1
errorcatch = False

testname = Sheet11.Cells(testrow, 3)

If testname = name Then
test = True
studentrow = testrow
errorcatch = True
Else
'do nothing
End If


If i = 10 Then
test = True
Else
'do nothing
End If
testrow = testrow + 1
Loop Until test = True



    If errorcatch = False Then
        studentrow = 201
        If suppresserror = False Then
            errormsg = "There was an error finding " & name & " in period " & period & ". please change the period and try again. "
             MsgBox errormsg
             Else
            'do nothing
            End If

    Else
    'do nothing
    End If

findstudentrow = studentrow
End Function

Public Sub placestudentinrow(placementrow As Integer, Optional entry As Integer = 1, Optional blankhours As Boolean = False)
Dim bufferrowstart As Integer
Dim entryrow As Integer
Dim templaterow As Integer
Dim replacecell As Boolean
Dim entryperiod As Integer


'this sets the class period that will be entered
If placementrow > 73 Then
entryperiod = 8
ElseIf placementrow > 63 Then
entryperiod = 7
ElseIf placementrow > 53 Then
entryperiod = 6
ElseIf placementrow > 43 Then
entryperiod = 5
ElseIf placementrow > 33 Then
entryperiod = 4
ElseIf placementrow > 23 Then
entryperiod = 3
ElseIf placementrow > 13 Then
entryperiod = 2
Else
entryperiod = 1
End If


MsgBox (entryperiod)




bufferrowstart = ((entry - 1) * 6) + 13
entryrow = bufferrowstart
templaterow = 1


'place gradebook-roster
Sheet11.Cells(placementrow, 1) = entryperiod
For i = 2 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet11.Cells(placementrow, i) = Sheet20.Cells(entryrow, i).Value
Else
'do nothing
End If

Next i
templaterow = templaterow + 1
entryrow = entryrow + 1

' Enter the correct classperiod
Sheet11.Cells(placementrow, 1) = classperiod

'attendance



If blankhours = False Then

For i = 1 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet14.Cells(placementrow, i) = Sheet20.Cells(entryrow, i).Value
Else
'do nothing
End If
Next i
templaterow = templaterow + 1
entryrow = entryrow + 1

ElseIf blankhours = True Then

For i = 1 To 4
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet14.Cells(placementrow, i) = Sheet20.Cells(entryrow, i).Value
Else
'do nothing
End If
Next i

templaterow = templaterow + 1
entryrow = entryrow + 1

For i = 4 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet14.Cells(placementrow, i) = ""
Else
'do nothing
End If
Next i


End If




'grade calculations
For i = 1 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet15.Cells(placementrow, i) = Sheet20.Cells(entryrow, i).Value
Else
'do nothing
End If

Next i
templaterow = templaterow + 1
entryrow = entryrow + 1



'cumulative grades
For i = 1 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet16.Cells(placementrow, i) = Sheet20.Cells(entryrow, i).Value
Else
'do nothing
End If

Next i
templaterow = templaterow + 1
entryrow = entryrow + 1



'weekly grades
For i = 1 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet17.Cells(placementrow, i) = Sheet20.Cells(entryrow, i).Value
Else
'do nothing
End If

Next i
templaterow = templaterow + 1
entryrow = entryrow + 1



'downgrades
For i = 1 To 150
replacecell = IsEmpty(Sheet20.Cells(templaterow, i))
If replacecell = True Then
Sheet18.Cells(placementrow, i) = Sheet20.Cells(entryrow, i)
Else
'do nothing
End If

Next i
templaterow = templaterow + 1
entryrow = entryrow + 1





End Sub

Sub clearrow(studentrow As Integer)

Call placestudentinrow(studentrow, -1)


End Sub

Private Sub CommandButton4_Click()
Dim note As String
Dim name As String

name = Sheet1.Cells(23, 2).Value
note = Sheet1.Cells(25, 1).Value


Call removestudent(name, note)

End Sub

Private Sub CommandButton5_Click()
'## move student button

Dim movefrom As Integer
Dim moveto As Integer
Dim name As String

movefrom = Sheet1.Cells(25, 2)
moveto = Sheet1.Cells(27, 2)
name = Sheet1.Cells(23, 2)


Call movestudent(name, movefrom, moveto)



End Sub
Public Sub movestudent(name As String, movefrom As Integer, moveto As Integer)
Dim oldrow As Integer
Dim newrow As Integer

'find row by name
oldrow = findstudentrow(name, movefrom)

'find slot
newrow = findslot(moveto)
'move to backup
Call placeintobackup(oldrow, 1)
'place in new slot slot
Sheet20.Cells(13, 1) = moveto
Call placestudentinrow(newrow)
Call clearrow(oldrow)


End Sub

Private Sub CommandButton6_Click()
'################### Split student button ##########################

Dim name As String
Dim fromperiod As Integer
Dim firstsplit As Integer
Dim secondsplit As Integer
Dim secondtest As Boolean
secondtest = False


name = Sheet1.Cells(23, 2)
fromperiod = Sheet1.Cells(25, 3)
firstsplit = Sheet1.Cells(27, 3)
secondsplit = Sheet1.Cells(28, 3)
secondtest = IsEmpty(Sheet1.Cells(28, 3))

If secondtest = True Then
Call splitstudent(name, fromperiod, firstsplit, secondsplit)
Else
Call splitstudent(name, fromperiod, firstsplit)
End If

Sheet1.Cells(23, 2) = ""
Sheet1.Cells(25, 3) = ""
Sheet1.Cells(27, 3) = ""
Sheet1.Cells(28, 3) = ""

End Sub

Public Sub splitstudent(name As String, fromperiod As Integer, Optional first As Integer = 0, Optional second As Integer = 0)
Dim callpro As Boolean
Dim slot As Integer



Call movestudent(name, fromperiod, fromperiod)

For i = 1 To 150
Sheet20.Cells(14, i) = ""
Next i



If first > 0 Then
slot = findslot(first)
Sheet20.Cells(13, 1) = first
Call placestudentinrow(slot)
Else
'do nothing
End If



If second > 0 Then
slot = findslot(second)
Sheet20.Cells(13, 1) = second
Call placestudentinrow(slot)
Else
'do nothing
End If


End Sub

Private Sub CommandButton7_Click()
'########################## This is the Combine Periods Button #################################

Dim firstC As Integer
Dim secondC As Integer
Dim thirdC As Integer
Dim fourthC As Integer
Dim name As String
Dim destination As Integer

firstC = Sheet1.Cells(25, 4)
secondC = Sheet1.Cells(26, 4)
thirdC = Sheet1.Cells(27, 4)
fourthC = Sheet1.Cells(28, 4)
name = Sheet1.Cells(23, 2)
destination = Sheet1.Cells(30, 4)




If fourthC > 0 Then
Call mergestudent(name, destination, firstC, secondC, thirdC, fourthC)
ElseIf thirdC > 0 Then
Call mergestudent(name, destination, firstC, secondC, thirdC)
ElseIf secondC > 0 Then
Call mergestudent(name, destination, firstC, secondC)
Else
MsgBox "There was an error, please check the periods and try again"
End If




End Sub

Public Sub mergestudent(name As String, destination As Integer, firstC As Integer, secondC As Integer, Optional thirdC As Integer = 0, Optional fourthC As Integer = 0)
Dim slot As Integer
Dim currentperiod As Integer
Dim backupnumber As Integer
Dim textaddendum As String
Dim hours As Integer
Dim count As Integer

hours = 0


textaddendum = "mergesave"
backupnumber = 0

'first to merge

currentperiod = firstC

If currentperiod > 0 Then
backupnumber = backupnumber + 1
slot = findstudentrow(name, currentperiod)
Call placeintobackup(slot, backupnumber)
Call clearrow(slot)
Else
'do nothing
End If



'second to merge
currentperiod = secondC

If currentperiod > 0 Then
backupnumber = backupnumber + 1
slot = findstudentrow(name, currentperiod)
Call placeintobackup(slot, backupnumber)
Call clearrow(slot)
Else
'do nothing
End If


'third to merge

currentperiod = thirdC

If currentperiod > 0 Then
backupnumber = backupnumber + 1
slot = findstudentrow(name, currentperiod)
Call placeintobackup(slot, backupnumber)
Call clearrow(slot)
Else
'do nothing
End If

'fourth to merge

currentperiod = fourthC

If currentperiod > 0 Then
backupnumber = backupnumber + 1
slot = findstudentrow(name, currentperiod)
Call placeintobackup(slot, backupnumber)
Call clearrow(slot)
Else
'do nothing
End If

'move to long term storage in order to save the information that would be lost through the merge.

Call placeintolongtermstorage(backupnumber, "this info is to make sure data is not lost", textaddendum)


'add the hours

count = 2
'we want to reach backup-1
Do
fhours = count * 6 + 8
hours = hours + Sheet20.Cells(fhours, 4).Value
count = count + 1
Loop Until count > (backupnumber)





hours = hours + Sheet20.Cells(14, 5)

Sheet20.Cells(14, 5) = hours

slot = findslot(destination)

Call placestudentinrow(slot)

'Call addextrahours(name, hours, destination)



End Sub


Private Sub CommandButton9_Click()
Dim classperiod(1 To 4) As Integer
Dim col As Integer
Dim newrow As Integer




classperiod(1) = Sheet1.Cells(15, 2).Value
classperiod(2) = Sheet1.Cells(15, 3).Value
classperiod(3) = Sheet1.Cells(15, 4).Value
classperiod(4) = Sheet1.Cells(15, 5).Value

Call newstudenttobackup

For i = 1 To 4
col = 1 + i
If IsEmpty(Sheet1.Cells(15, col)) = False Then
newrow = findslot(classperiod(i))
Call placestudentinrow(newrow)
Else
'do nothing
End If

Next i



'this clears the new student template
Sheet1.Cells(15, 2) = ""
Sheet1.Cells(15, 3) = ""
Sheet1.Cells(15, 4) = ""
Sheet1.Cells(15, 5) = ""
Sheet1.Cells(14, 2) = ""
Sheet1.Cells(16, 2) = ""
Sheet1.Cells(17, 2) = ""
Sheet1.Cells(18, 2) = ""

End Sub


Public Sub placeallhours(name As String)

Dim test(1 To 8) As Integer
Dim classes(1 To 8)

For i = 1 To 8



Next i




End Sub

Public Sub recordandresetstudent(name As String)

Dim slot As Integer
Dim entry As Integer
Dim period As Integer

For i = 1 To 8
period = i
slot = findstudentrow(name, period, True)

If slot = 201 Then

'do nothing
Else
entry = entry + 1
Call placeintobackup(slot, entry)

Sheet11.Cells(slot, 77) = Sheet11.Cells(slot, 76).Value - Sheet11.Cells(slot, 75).Value


Sheet15.Cells(slot, 5) = Sheet15.Cells(slot, 3)
Sheet15.Cells(slot, 4) = Sheet15.Cells(slot, 6)





For u = 1 To 54
Sheet11.Cells(slot, (u + 19)) = ""
Next u



End If




Next i


Dim note As String
Dim fileaddendum As String

note = "placed into storage to save a record when grades were reset."
fileaddendum = "-resetsave"
Call placeintolongtermstorage(entry, note, fileaddendum)


End Sub
