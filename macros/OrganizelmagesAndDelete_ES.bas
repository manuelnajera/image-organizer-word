Attribute VB_Name = "Module1"
Sub OrganizeImagesAndDelete()
    Dim wdDoc As Document
    Dim wdRng As Range
    Dim imgPath As String
    Dim layout As Integer
    Dim imgFileName As String
    Dim imagesPerLine As Integer
    Dim imagesPerPage As Integer
    Dim currentPageWidth As Single
    Dim currentPageHeight As Single
    Dim imgWidth As Double
    Dim imgHeight As Double
    Dim imgMargin As Double ' Margin between images in points
    
    ' Set the layout option
    layout = InputBox("Imágenes por página (1, 2, 4, 6, 8, 10, hasta 20):")
    If layout < 1 Or layout > 20 Then
        ' Show an alert for invalid value
        MsgBox "Valor inválido! Por favor introduzca un valor entre 1 y 20.", vbExclamation
        Exit Sub ' Exit the subroutine
    End If
    
    If layout > 2 Then
    imagesPerLine = InputBox("Imágenes por fila  (1, 2, 3, 4, 5):")
    Else
    imagesPerLine = 1
    End If
    
    If imagesPerLine < 1 Or imagesPerLine > 5 Then
        ' Show an alert for invalid value
        MsgBox "Valor inválido! Por favor introduzca un valor entre  1 y 5.", vbExclamation
        Exit Sub ' Exit the subroutine
    End If
    
    imagesPerPage = layout ' Store the original layout value
    
    ' Create a new Word document
    Set wdDoc = Documents.Add
    Set wdRng = wdDoc.Content
    
    With wdDoc.PageSetup
        .LeftMargin = InchesToPoints(0.5) ' 0.5 inches left margin
        .RightMargin = InchesToPoints(0.5) ' 0.5 inches right margin
        .TopMargin = InchesToPoints(0.5) ' 0.5 inches top margin
        .BottomMargin = InchesToPoints(0.5) ' 0.5 inches bottom margin
    End With
    
    ' Set the margin between images (adjust as needed)
    imgMargin = InchesToPoints(0.3) ' 0.3 inches margin
    
    ' Hard-code the page size to letter size (8.5 inches by 11 inches)
    currentPageWidth = 7.5 * 72 ' 8.5 inches converted to points (1 inch = 72 points)
    currentPageHeight = 10 * 72 ' 11 inches converted to points (1 inch = 72 points)
    
    ' Loop through files in the folder
    imgPath = ThisDocument.Path & "/"
    imgFileName = Dir(imgPath)
    Do While imgFileName <> ""
        ' Check if the file is an image
        If IsImageFile(imgFileName) Then
            ' Insert the image
            wdRng.InlineShapes.AddPicture fileName:=imgPath & imgFileName, LinkToFile:=False, SaveWithDocument:=True
            wdRng.InsertBefore " "
            ' Delete the file
           Kill imgPath & imgFileName
            
' Check if there are InlineShapes in the range
If wdRng.InlineShapes.Count > 0 Then
    Dim horizontalPrefered As Boolean
    ' Calculate dimensions for each image
    imgWidth = (currentPageWidth - (imagesPerLine - 1) * imgMargin) / imagesPerLine
    imgHeight = (currentPageHeight / (layout / imagesPerLine)) - imgMargin
                    
    If imgWidth > imgHeight Then
        horizontalPrefered = True
    Else
        horizontalPrefered = False
    End If

    ' Resize the first image
    With wdRng.InlineShapes(1)
        .LockAspectRatio = msoTrue ' Preserve aspect ratio
                    
        If horizontalPrefered Then
            If .height > .width Then
                .height = imgWidth
                If .width > imgHeight Then
                    .width = imgHeight
                End If
                ' Convert the InlineShape to a Shape
            RotateInlineShape wdRng.InlineShapes(1), 90
            Else
                .width = imgWidth
                    If .height > imgHeight Then
                    .height = imgHeight
                End If
            End If
        Else
            If .width > .height Then
                .width = imgHeight
                If .height > imgWidth Then
                    .height = imgWidth
                End If
                ' Convert the InlineShape to a Shape
            RotateInlineShape wdRng.InlineShapes(1), 90
            Else
                .height = imgHeight
                    If .width > imgWidth Then
                    .width = imgWidth
                End If
            End If
        End If
    End With

 
    ' Insert a space (margin) between images
    wdRng.InsertBefore "  " ' Insert a space


                Else
                    MsgBox "No se encontraron im�genes."
                End If

         End If
        ' Move to the next file
        imgFileName = Dir
    Loop
    ConvertShapesToInlineShapes
    ' Save the document
    wdDoc.SaveAs2 fileName:=ThisDocument.Path & "/ImagenesOrganizadas.docx"
    
    ' Close the document
    wdDoc.Activate
    
    ' Clean up
    Set wdRng = Nothing
    Set wdDoc = Nothing
End Sub

Function IsImageFile(ByVal fileName As String) As Boolean
    ' Check if the file has an image extension
    Select Case LCase(Right(fileName, Len(fileName) - InStrRev(fileName, ".")))
        Case "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff"
            IsImageFile = True
        Case Else
            IsImageFile = False
    End Select
End Function

Function ConvertShapesToInlineShapes()
    Dim shp As Shape
    
    ' Loop through each shape in the document
    For Each shp In ActiveDocument.Shapes
        ' Check if the shape is not an InlineShape
        If Not shp.Type = msoInlineShape Then
            ' Convert the shape to InlineShape
            shp.ConvertToInlineShape
        End If
    Next shp
End Function
Function RotateInlineShape(ByVal inlineShape As inlineShape, ByVal degrees As Integer)
    Dim shp As Shape
    Set shp = inlineShape.ConvertToShape
    
    ' Rotate the shape by the specified degrees
    shp.Rotation = degrees
    
    ' Convert the shape back to InlineShape
    shp.ConvertToInlineShape
End Function

