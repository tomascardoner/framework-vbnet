Namespace CardonerSistemas

    Module Barcodes
        Friend Const QRCodeDataField As String = "{DATA}"

        Friend Function CreateQR() As Bitmap

            'Create a new QR bitmap image
            Dim bmp As New Bitmap(21, 21)

            'Get the graphics object to manipulate the bitmap
            Dim gr As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)

            'Set the background of the bitmap to white
            gr.FillRectangle(Brushes.White, 0, 0, 21, 21)

            'Draw position detection patterns
            'Top Left
            gr.DrawRectangle(Pens.Black, 0, 0, 6, 6)
            gr.FillRectangle(Brushes.Black, 2, 2, 3, 3)

            'Top Right
            gr.DrawRectangle(Pens.Black, 14, 0, 6, 6)
            gr.FillRectangle(Brushes.Black, 2, 16, 3, 3)

            'Bottom Left
            gr.DrawRectangle(Pens.Black, 0, 14, 6, 6)
            gr.FillRectangle(Brushes.Black, 16, 2, 3, 3)


            '*** Drawing pixels is done off the bitmap object, not the graphics object

            'Arbitrary black pixel
            bmp.SetPixel(8, 14, Color.Black)

            'Top timing pattern
            bmp.SetPixel(8, 6, Color.Black)
            bmp.SetPixel(10, 6, Color.Black)
            bmp.SetPixel(12, 6, Color.Black)

            'Left timing pattern
            bmp.SetPixel(6, 8, Color.Black)
            bmp.SetPixel(6, 10, Color.Black)
            bmp.SetPixel(6, 12, Color.Black)

            'Add code here to set the rest of the pixels as needed
        End Function

    End Module

End Namespace