PPT-Clone
=========

PPT Slide Creation in VBA
Public slideNum As Variant
Public tabNum As Variant

Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long

Public Function ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Function


Sub ExcelToNewPowerPoint()
    Dim PPApp As PowerPoint.Application
    Dim PPPres As PowerPoint.Presentation
    Dim PPSlide As PowerPoint.Slide
    Dim sld1 As PowerPoint.Slide
    
    ' Create instance of PowerPoint
    Set PPApp = CreateObject("Powerpoint.Application")

    ' For automation to work, PowerPoint must be visible
    ' (alternatively, other extraordinary measures must be taken)
    PPApp.Visible = True
    PPApp.Presentations.Open ("C:\Users\tpei\Documents\Commercial Finance\CF Metrics PPT\PPT\Q3FY13 Feb - CF Metrics.pptx")
    ' Create a presentation
    'Set PPPres = PPApp.Presentations.Add

    ' Some PowerPoint actions work best in normal slide view
    PPApp.ActiveWindow.ViewType = ppViewSlide

    ' Reference active presentation
    Set PPPres = PPApp.ActivePresentation
    PPApp.ActiveWindow.ViewType = ppViewSlide
    'PowerPoint.Application.ScreenUpdating = False
    
    For x = 2 To 116 'Worksheets.Count
    
    On Error GoTo errhandler:
    
    Range("a1").Activate
    Worksheets(x).Select
    Worksheets(x).Activate
    Range("a1").Activate
    slideNum = Range("a60").Value
    Debug.Print x & Worksheets(x).Name
    
   'If x <> 77 Then
   
           If slideNum < 155 Then
        
            Set PPSlide = PPPres.Slides(slideNum)
            
            Select Case Range("A62").Value
        
                Case 1
        
                   ActiveSheet.ChartObjects(1).Activate
        
                    ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, _
                     Format:=xlPicture
                
                     PPApp.ActiveWindow.View.GotoSlide slideNum
          
                     PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue, , , "test").Select
                        PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.99, msoFalse, msoScaleFromTopLeft
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 45
        
                        'paste in table at the bottom
                                 
                        Worksheets(x).Activate
                       ActiveSheet.Range("A9:I17").Select
                       Selection.Copy
                        
                        PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                        PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.85, msoFalse, msoScaleFromTopLeft
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 372
        
        
        
                Case 2
                        
                                   ActiveSheet.ChartObjects(1).Activate
                        
                          ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, _
                                Format:=xlPicture
                                PPApp.ActiveWindow.View.GotoSlide slideNum
                          
                            PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                            
                        PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.95, msoFalse, msoScaleFromTopLeft
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 45
                        
                        'paste in table at the bottom
                        ActiveSheet.Range("A9:I17").Select
                            Selection.Copy
                                
                         PPApp.ActiveWindow.View.GotoSlide slideNum
                          
                            PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                                
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 372
        
                Case 3
        
                                  
                        ActiveSheet.ChartObjects(1).Activate
                        
                          ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, _
                                Format:=xlPicture
                                    
                         PPApp.ActiveWindow.View.GotoSlide slideNum
                          
                            PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                                
                        PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 1, msoFalse, msoScaleFromTopLeft
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 45
                        
                        'paste in table at the bottom
                        ActiveSheet.Range("N1:T8").Select
                            Selection.Copy
                    
                   
                         PPApp.ActiveWindow.View.GotoSlide slideNum
                          
                           PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                            
                       PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 372
                        
        
                Case 4
        
                               ActiveSheet.ChartObjects(1).Activate
                    
                      ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, _
                            Format:=xlPicture
                            
                     PPApp.ActiveWindow.View.GotoSlide slideNum
                      
                        PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                        
                    PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.95, msoFalse, msoScaleFromTopLeft
                    PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                    PPApp.ActiveWindow.Selection.ShapeRange.Top = 80
        
        
                Case 5
                        ActiveSheet.Range("B5:R43").Select
                            Selection.Copy
                            
                                    
                         PPApp.ActiveWindow.View.GotoSlide slideNum
                          
                            PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                                
                          PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.76, msoFalse, msoScaleFromTopLeft
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 53
        
                 
                Case 6
                        
                        ActiveSheet.Range("B2:M20").Select
                            Selection.Copy
                            
                                    
                         PPApp.ActiveWindow.View.GotoSlide slideNum
                          
                            PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue).Select
                                
                        PPApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.88, msoFalse, msoScaleFromTopLeft
                        PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
                        PPApp.ActiveWindow.Selection.ShapeRange.Top = 130
        
                Case Else
        
                    'do nothing
                    Range("a1").Value = Range("a1").Value
        
                End Select
             
             End If
    
  ' End If
   
    
  Next
        
    ' Reference active slide
   ' Set PPSlide = PPPres.Slides(slideNum) ' _
   ' (PPApp.ActiveWindow.Selection.SlideRange.SlideIndex)
    
    ' Copy chart as a picture
   'ThisWorkbook.Worksheets("Sheet1").Activate
'ActiveSheet.ChartObjects(1).Activate
'ActiveSheet.Range("B2:M20").Select
    'Selection.Copy
 
  'ActiveChart.CopyPicture Appearance:=xlScreen, Size:=xlScreen, _
        Format:=xlPicture

        

    ' Paste chart
    
 ' PPApp.ActiveWindow.View.GotoSlide 8
  
   ' PPSlide.Shapes.PasteSpecial(ppPasteMetafilePicture, msoTrue, , , "test").Select
     
    ' Align pasted chart
   ' PPApp.ActiveWindow.Selection.ShapeRange.Left = 100
   ' PPApp.ActiveWindow.Selection.ShapeRange.Top = 380
    
  '  PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, True
    'PPApp.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, True







    ' Save and close presentation
   With PPPres
       .Save
       .Close
   End With

    ' Quit PowerPoint
    PPApp.Quit

    ' Clean up
    Set PPSlide = Nothing
    Set PPPres = Nothing
    Set PPApp = Nothing

Exit Sub
errhandler:
  ClearClipboard
     With PPPres
       .Save
       
   End With

    ' Quit PowerPoint
   ' PPApp.Quit
  
Resume Next


End Sub
