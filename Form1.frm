VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuCreateAccessContentDatabase 
         Caption         =   "Create &Access Content Database"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim db As Database
Dim tblDef As TableDef
Dim fldDef As Field
Dim fldLoop As Field
Dim prpLoop As Property
Dim indx As Index
Dim dbName As String
Dim FileNumber As String
Dim FileName As String



Private Sub mnuCreateAccessContentDatabase_Click()
'Replace with common dialog code
dbName = "D:\Alex\VB\Create Database\test.mdb"

If (Len(Dir(dbName))) Then
    Kill dbName
End If

Set db = DBEngine.Workspaces(0).CreateDatabase(dbName, dbLangGeneral)

Call createCustomer_Tracking
Call createCommentTracking
Call createCourse
Call createELO
Call createGraphic
Call createGraphicData
Call createGraphicItem
Call createGraphicItemType
Call createGraphicItemValue
Call createGraphicReference
Call createGraphicTracking
Call createLesson
Call createLessonTracking
Call createMediaType
Call createPage
Call createPageItem
Call createPageItemType
Call createPageType
Call createPersonnel
Call createQuestion
Call createUnit
'Call createRelationships

MsgBox ("Database successfully created at " & dbName & "." & _
        vbCrLf & "Database Info file created at " & FileName & ".")

End Sub

Public Sub createCustomer_Tracking()
Set tblDef = db.CreateTableDef("Customer_Tracking")

With tblDef
    .Fields.Append .CreateField("Customer_ID", dbLong)
    .Attributes = dbAutoIncrField
    .Fields.Append .CreateField("Customer_Num", dbText, 4)
    .Fields.Append .CreateField("Title", dbText, 64)
    .Fields.Append .CreateField("Action", dbText, 32)
    .Fields.Append .CreateField("Frequency", dbText, 32)
    .Fields.Append .CreateField("DateSubmitted", dbDate)
    .Fields.Append .CreateField("DateDue", dbDate)
    .Fields.Append .CreateField("Responsibility", dbLong)
End With

db.TableDefs.Append tblDef

Set indx = tblDef.CreateIndex("PrimaryKey")
Set fldDef = indx.CreateField("Customer_ID")
    With indx
        .Primary = True
        .Fields.Append fldDef
    End With
tblDef.Indexes.Append indx

FileNumber = FreeFile
FileName = "D:\Alex\VB\Create Database\Database Info.txt"

Open FileName For Output As #FileNumber
Print #FileNumber, vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createCommentTracking()
Set tblDef = db.CreateTableDef("CommentTracking")

With tblDef
    .Fields.Append .CreateField("CommentTrackingID", dbLong)
    .Attributes = dbAutoIncrField
    .Fields.Append .CreateField("LessonID", dbLong)
    .Fields.Append .CreateField("PageName", dbText, 32)
    .Fields.Append .CreateField("CommentDate", dbDate)
    .Fields.Append .CreateField("Reviewer", dbLong)
    .Fields.Append .CreateField("CommentText", dbMemo)
    .Fields.Append .CreateField("ResponseText", dbMemo)
    .Fields.Append .CreateField("ResponseDate", dbDate)
    .Fields.Append .CreateField("Responder", dbLong)
End With

db.TableDefs.Append tblDef

Set indx = tblDef.CreateIndex("PrimaryKey")
Set fldDef = indx.CreateField("CommentTrackingID")
    With indx
        .Primary = True
        .Fields.Append fldDef
    End With
tblDef.Indexes.Append indx

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop
    Next fldLoop

Close #FileNumber

End Sub

Public Sub createCourse()
Set tblDef = db.CreateTableDef("Course")

With tblDef
    .Fields.Append .CreateField("CourseID", dbLong)
    .Attributes = dbAutoIncrField
    .Fields.Append .CreateField("CourseCode", dbText, 16)
    .Fields.Append .CreateField("CourseTitle", dbText, 128)
End With

db.TableDefs.Append tblDef

Set indx = tblDef.CreateIndex("PrimaryKey")
Set fldDef = indx.CreateField("CourseID")
    With indx
        .Primary = True
        .Fields.Append fldDef
    End With
tblDef.Indexes.Append indx

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createELO()
Set tblDef = db.CreateTableDef("ELO")

With tblDef
    .Fields.Append .CreateField("ELO_ID", dbLong)
    .Fields.Append .CreateField("LessonID", dbLong)
    .Fields.Append .CreateField("ELO_Number", dbText, 50)
    .Fields.Append .CreateField("ELO_Title", dbText, 50)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createGraphic()
Set tblDef = db.CreateTableDef("Graphic")

With tblDef
    .Fields.Append .CreateField("GraphicID", dbLong)
    .Fields.Append .CreateField("MediaType", dbLong)
    .Fields.Append .CreateField("FileName", dbText, 64)
    .Fields.Append .CreateField("FileLocation", dbText, 128)
    .Fields.Append .CreateField("GraphicTitle", dbText, 128)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber
    
End Sub

Public Sub createGraphicData()
Set tblDef = db.CreateTableDef("GraphicData")

With tblDef
    .Fields.Append .CreateField("GraphicID", dbLong)
    .Fields.Append .CreateField("ProjectID", dbLong)
    .Fields.Append .CreateField("CourseID", dbLong)
    .Fields.Append .CreateField("Frame", dbText, 50)
    .Fields.Append .CreateField("Classification", dbInteger)
    .Fields.Append .CreateField("SystemCode", dbText, 50)
    .Fields.Append .CreateField("SubSystem", dbText, 50)
    .Fields.Append .CreateField("FileName", dbText, 50)
    .Fields.Append .CreateField("FileSize", dbText, 50)
    .Fields.Append .CreateField("FileDate", dbText, 50)
    .Fields.Append .CreateField("FileTime", dbText, 50)
    .Fields.Append .CreateField("FileType", dbText, 50)
    .Fields.Append .CreateField("FileLocation", dbText, 50)
    .Fields.Append .CreateField("MediaVolume", dbText, 50)
    .Fields.Append .CreateField("CDLocation", dbText, 50)
    .Fields.Append .CreateField("Description", dbText, 128)
    .Fields.Append .CreateField("OnLineCode", dbLong)
    .Fields.Append .CreateField("SourceOrFinal", dbText, 50)
    .Fields.Append .CreateField("Comments", dbText, 50)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createGraphicItem()
Set tblDef = db.CreateTableDef("GraphicItem")

With tblDef
    .Fields.Append .CreateField("GraphicID", dbLong)
    .Fields.Append .CreateField("GraphicItem", dbLong)
    .Fields.Append .CreateField("GraphicItemType", dbLong)
    .Fields.Append .CreateField("Range", dbLong)
    .Fields.Append .CreateField("DefaultValue", dbLong)
    .Fields.Append .CreateField("XLoc", dbLong)
    .Fields.Append .CreateField("YLoc", dbLong)
    .Fields.Append .CreateField("XLen", dbLong)
    .Fields.Append .CreateField("YLen", dbLong)
    .Fields.Append .CreateField("ItemTitle", dbText, 64)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createGraphicItemValue()
Set tblDef = db.CreateTableDef("GraphicItemValue")

With tblDef
    .Fields.Append .CreateField("GraphicItemTypeID", dbLong)
    .Fields.Append .CreateField("GraphicItemType", dbText, 16)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createGraphicReference()
Set tblDef = db.CreateTableDef("GraphicReference")

With tblDef
    .Fields.Append .CreateField("SectionID", dbLong)
    .Fields.Append .CreateField("PgeNum", dbLong)
    .Fields.Append .CreateField("PageItem", dbLong)
    .Fields.Append .CreateField("GraphicItem", dbLong)
    .Fields.Append .CreateField("GraphicItemValue", dbLong)
    .Fields.Append .CreateField("GraphicItemText", dbText, 128)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createGraphicTracking()
Set tblDef = db.CreateTableDef("GraphicTracking")

With tblDef
    .Fields.Append .CreateField("GraphicTrackingID", dbLong)
    .Fields.Append .CreateField("GraphicID", dbLong)
    .Fields.Append .CreateField("Action", dbText, 32)
    .Fields.Append .CreateField("PersonnelID", dbLong)
    .Fields.Append .CreateField("ActionDate", dbDate)
    .Fields.Append .CreateField("LastAction", dbLong)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createLesson()
Set tblDef = db.CreateTableDef("Lesson")

With tblDef
    .Fields.Append .CreateField("LessonID", dbLong)
    .Fields.Append .CreateField("UnitID", dbLong)
    .Fields.Append .CreateField("LessonTitle", dbText, 128)
    .Fields.Append .CreateField("LessonNumber", dbText, 32)
    .Fields.Append .CreateField("PrimaryJob", dbText, 32)
    .Fields.Append .CreateField("SecondaryJob", dbText, 32)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createLessonTracking()
Set tblDef = db.CreateTableDef("LessonTracking")

With tblDef
    .Fields.Append .CreateField("LessonTrackingID", dbLong)
    .Fields.Append .CreateField("Action", dbText, 50)
    .Fields.Append .CreateField("ActionDate", dbDate)
    .Fields.Append .CreateField("LastAction", dbLong)
    .Fields.Append .CreateField("Comments", dbMemo)
    .Fields.Append .CreateField("Value", dbInteger)
    .Fields.Append .CreateField("LessonTitle", dbText, 128)
    .Fields.Append .CreateField("LessonType", dbInteger)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createMediaType()
Set tblDef = db.CreateTableDef("MediaType")

With tblDef
    .Fields.Append .CreateField("MediaTypeID", dbLong)
    .Fields.Append .CreateField("MediaType", dbText, 16)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber
    
End Sub

Public Sub createPage()
Set tblDef = db.CreateTableDef("Page")

With tblDef
    .Fields.Append .CreateField("SectionID", dbLong)
    .Fields.Append .CreateField("PageName", dbText, 128)
    .Fields.Append .CreateField("PageNum", dbLong)
    .Fields.Append .CreateField("PageType", dbLong)
    .Fields.Append .CreateField("NextPage", dbLong)
    .Fields.Append .CreateField("PrevPage", dbLong)
    .Fields.Append .CreateField("Fault", dbText, 32)
    .Fields.Append .CreateField("JobNext", dbLong)
    .Fields.Append .CreateField("JobBack", dbLong)
    .Fields.Append .CreateField("StartFlag", dbInteger)
    .Fields.Append .CreateField("AskTheExpert", dbLong)
    .Fields.Append .CreateField("Functional", dbLong)
    .Fields.Append .CreateField("Schematic", dbLong)
    .Fields.Append .CreateField("ThreeDObject", dbLong)
    .Fields.Append .CreateField("Layer", dbLong)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createPageItem()
Set tblDef = db.CreateTableDef("PageItem")

With tblDef
    .Fields.Append .CreateField("SectionID", dbLong)
    .Fields.Append .CreateField("PageNum", dbLong)
    .Fields.Append .CreateField("SubPage", dbLong)
    .Fields.Append .CreateField("PageItem", dbLong)
    .Fields.Append .CreateField("PageItemType", dbLong)
    .Fields.Append .CreateField("XLoc", dbLong)
    .Fields.Append .CreateField("YLoc", dbLong)
    .Fields.Append .CreateField("XLen", dbLong)
    .Fields.Append .CreateField("YLen", dbLong)
    .Fields.Append .CreateField("Label", dbMemo)
    .Fields.Append .CreateField("GraphicID", dbLong)
    .Fields.Append .CreateField("Layer", dbLong)
    .Fields.Append .CreateField("DisplayMode", dbText, 32)
    .Fields.Append .CreateField("Judge", dbText, 32)
    .Fields.Append .CreateField("Logic", dbText, 128)
    .Fields.Append .CreateField("Options", dbText, 128)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createPageItemType()
Set tblDef = db.CreateTableDef("PageItemType")

With tblDef
    .Fields.Append .CreateField("PageItemTypeID", dbLong)
    .Fields.Append .CreateField("PageItemType", dbText, 16)
End With

db.TableDefs.Append tblDef
    
FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createPageType()
Set tblDef = db.CreateTableDef("PageType")

With tblDef
    .Fields.Append .CreateField("PageTypeID", dbLong)
    .Fields.Append .CreateField("PageType", dbText, 50)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createPersonnel()
Set tblDef = db.CreateTableDef("Personnel")

With tblDef
    .Fields.Append .CreateField("PersonnelID", dbLong)
    .Fields.Append .CreateField("LastName", dbText, 32)
    .Fields.Append .CreateField("FirstName", dbText, 32)
    .Fields.Append .CreateField("Title", dbText, 32)
    .Fields.Append .CreateField("DevelopmentTeam", dbText, 50)
    .Fields.Append .CreateField("JobFunction", dbText, 32)
    .Fields.Append .CreateField("PhoneNumber", dbText, 32)
    .Fields.Append .CreateField("EMailAddress", dbText, 64)
    .Fields.Append .CreateField("AccessClass", dbText, 32)
    .Fields.Append .CreateField("LoginName", dbText, 32)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createQuestion()
Set tblDef = db.CreateTableDef("Question")

With tblDef
    .Fields.Append .CreateField("ELO_ID", dbLong)
    .Fields.Append .CreateField("QNumber", dbLong)
    .Fields.Append .CreateField("QBank", dbText, 32)
    .Fields.Append .CreateField("QType", dbLong)
    .Fields.Append .CreateField("Stem", dbText, 128)
    .Fields.Append .CreateField("CorrectAnswer", dbText, 128)
    .Fields.Append .CreateField("Distractor1", dbText, 128)
    .Fields.Append .CreateField("Distractor2", dbText, 128)
    .Fields.Append .CreateField("Distractor3", dbText, 128)
    .Fields.Append .CreateField("Distractor4", dbText, 200)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createUnit()
Set tblDef = db.CreateTableDef("Unit")

With tblDef
    .Fields.Append .CreateField("UnitID", dbLong)
    .Fields.Append .CreateField("CourseID", dbLong)
    .Fields.Append .CreateField("UnitNumber", dbText, 32)
    .Fields.Append .CreateField("UnitTitle", dbText, 128)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub

Public Sub createRelationships()

End Sub

Public Sub createGraphicItemType()
Set tblDef = db.CreateTableDef("GraphicItemType")

With tblDef
    .Fields.Append .CreateField("GraphicItemTypeID", dbLong)
    .Fields.Append .CreateField("GraphicItemType", dbText, 16)
End With

db.TableDefs.Append tblDef

FileNumber = FreeFile

Open FileName For Append As #FileNumber
Print #FileNumber, vbCrLf & vbCrLf & "Properties of Fields in " & tblDef.Name & vbCrLf
    For Each fldLoop In tblDef.Fields
        Print #FileNumber, "    " & fldLoop.Name
        For Each prpLoop In fldLoop.Properties
            On Error Resume Next
            Print #FileNumber, "        " & prpLoop.Name & " - " & _
                IIf(prpLoop = "", "[empty]", prpLoop)
            On Error GoTo 0
        Next prpLoop

    Next fldLoop
 Close #FileNumber

End Sub
