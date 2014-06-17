Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO

Public Class TiffImageSplitter
    Public Function getPageCount(ByVal fileName As String) As Integer
        Dim pageCount = -1
        Try
            Dim img = Bitmap.FromFile(fileName)
            pageCount = img.GetFrameCount(FrameDimension.Page)
            img.Dispose()
        Catch ex As Exception
            pageCount = 0
        End Try

        Return pageCount
    End Function

    'Retrieve a specific Page from a multi-page tiff image
    Public Function getTiffImage(ByVal sourceFile As String, ByVal pageNumber As Integer) As Image
        Dim returnImage As Image
        Try
            Dim sourceImage = Bitmap.FromFile(sourceFile)
            returnImage = getTiffImage(sourceImage, pageNumber)
            sourceImage.Dispose()
        Catch ex As Exception
            returnImage = Nothing
        End Try
        Return returnImage
    End Function

    Public Function getTiffImage(ByVal sourceImage As Image, ByVal pageNumber As Integer) As Image
        Dim ms As MemoryStream
        Dim returnImage As Image
        Try
            ms = New MemoryStream
            Dim objGuid = sourceImage.FrameDimensionsList(0)
            Dim objDimension = New FrameDimension(objGuid)
            sourceImage.SelectActiveFrame(objDimension, pageNumber)
            sourceImage.Save(ms, ImageFormat.Tiff)
            returnImage = Image.FromStream(ms)
        Catch ex As Exception
            returnImage = Nothing
        End Try
        Return returnImage
    End Function

End Class
