Imports System.IO
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO
Imports PdfSharp.Drawing
Imports PdfSharp.Drawing.XImage

Module Module1
    Dim tiff As New TiffImageSplitter

    Sub Main()
        Console.WriteLine("Enter path and name to list file: ")
        Dim fileList = Console.ReadLine()
        Dim line As String
        Dim lineArr() As String
        Dim pdf As New PdfDocument
        Dim pdf2 As New PdfDocument
        Try
            If File.Exists(fileList) Then
                Dim sr As New StreamReader(fileList)
                Do While sr.Peek() <> -1
                    line = sr.ReadLine
                    lineArr = line.Split("|")
                    'Console.WriteLine(lineArr(0))
                    'Console.WriteLine(lineArr(1))
                    Console.WriteLine(lineArr(2))

                    Dim pageCount = tiff.getPageCount(lineArr(1))

                    For i As Integer = 0 To pageCount - 1
                        Dim page As New PdfPage
                        Dim tiffImg As System.Drawing.Image
                        tiffImg = tiff.getTiffImage(lineArr(1), i)
                        Dim img = XImage.FromGdiPlusImage(tiffImg)

                        page.Width = img.PointWidth
                        page.Height = img.PointHeight
                        pdf.Pages.Add(page)

                        Dim xgr = XGraphics.FromPdfPage(pdf.Pages(i))

                        xgr.DrawImage(img, 0, 0)
                        Console.WriteLine("Page #" + (i + 1).ToString)
                    Next

                    pdf.Save(lineArr(2))
                    pdf.Close()

                Loop
            Else
                Console.WriteLine("File didn't exist!")
            End If
            'Console.ReadLine()
        Catch ex As Exception

        End Try
    End Sub

End Module
