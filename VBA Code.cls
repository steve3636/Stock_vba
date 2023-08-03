VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Analysis()

 

    Dim Last_row As Long

    Dim First_ticker As Long

    Dim Next_ticker As Long

    Dim Ticker As Long

    Dim Ticker_Id As String

    Dim NextTicker_Id As String

    Dim Volume As Double

    Dim Year_close As Double

    Dim Year_open As Double

    Dim Tab_num As Integer

    Dim ws As Worksheet

   

    Tab_num = ThisWorkbook.Sheets.Count

   

    For j = 1 To Tab_num

        Set ws = ThisWorkbook.Sheets(j)

        ws.Activate

       

        Ticker = 1

        First_ticker = 2

       

        ws.Range("J1").Value = "Ticker"

        ws.Range("K1").Value = "Yearly Change"

        ws.Range("L1").Value = "Percent Change"

        ws.Range("M1").Value = "Total Stock Volume"

       

        ws.Range("P1").Value = "Ticker"

        ws.Range("Q1").Value = "Value"

        ws.Range("O2").Value = "Greatest % Increase"

        ws.Range("O3").Value = "Greatest % Decrease"

        ws.Range("O4").Value = "Greatest Total Valume"

              

               

        Last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

       

        Volume = ws.Cells(2, 7).Value

       

        For i = 2 To Last_row

            Ticker_Id = ws.Cells(i, 1).Value

            NextTicker_Id = ws.Cells(i + 1, 1).Value

           

            If Ticker_Id = NextTicker_Id Then

                Volume = Volume + ws.Cells(i + 1, 7).Value

            Else

                Ticker = Ticker + 1

                Next_ticker = i + 1

                ws.Cells(Ticker, 10).Value = Ticker_Id

                ws.Cells(Ticker, 13).Value = Volume

               

                Year_close = ws.Cells(Next_ticker - 1, 6).Value

                Year_open = ws.Cells(First_ticker, 3).Value

               

                While Year_open = 0 And First_ticker < Next_ticker

                    First_ticker = First_ticker + 1

                    Year_open = ws.Cells(First_ticker, 3).Value

                Wend

               

                ws.Cells(Ticker, 11).Value = Year_close - Year_open

                ws.Cells(Ticker, 12).Value = (Year_close - Year_open) / Year_open

               

                ws.Cells(Ticker, 12).NumberFormat = "0.00%"

               

                If ws.Cells(Ticker, 11).Value > 0 Then

                    ws.Cells(Ticker, 11).Interior.ColorIndex = 4

                Else

                    ws.Cells(Ticker, 11).Interior.ColorIndex = 3

                End If

               

                First_ticker = Next_ticker

                Volume = ws.Cells(First_ticker, 7).Value

            End If

        Next i

       

        ws.Range("J1:M1").Columns.AutoFit

    Next j

   

End Sub


