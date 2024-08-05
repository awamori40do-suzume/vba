Sub GenerateAndSaveQRCode()
    Dim ws As Worksheet
    Dim textToEncode As String
    Dim barcode As OLEObject
    Dim cell As Range
    Dim imagePath As String

    ' シートを設定
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' カレントディレクトリのパスを取得
    imagePath = ThisWorkbook.Path & "\"

    ' A1からA10の範囲の各セルをループ
    For Each cell In ws.Range("A1:A10")
        textToEncode = cell.Value

        If textToEncode <> "" Then
            ' 新しいバーコードコントロールを作成
            Set barcode = ws.OLEObjects.Add(ClassType:="BARCODE.BarCodeCtrl.16", _
                                             Link:=False, _
                                             DisplayAsIcon:=False, _
                                             Left:=cell.Offset(0, 1).Left, _
                                             Top:=cell.Top, _
                                             Width:=100, _
                                             Height:=100)

            ' バーコードコントロールのプロパティを設定
            With barcode.Object
                .Style = 11 ' QR Code
                .Value = textToEncode
                .ShowText = False
                .AutoSize = True
            End With

            ' バーコードをシェイプとしてコピー
            barcode.Copy
            ws.Pictures.Paste.Select

            ' コピーしたシェイプを選択して保存
            With Selection
                .Name = "QRCode"
                .CopyPicture
                SavePictureToFile imagePath & "QRCode_" & cell.Row & ".png"
                .Delete
            End With

            ' バーコードコントロールを削除
            barcode.Delete
        End If
    Next cell

    MsgBox "QRコードの生成と保存が完了しました。"
End Sub

Sub SavePictureToFile(filePath As String)
    ' 画像として保存するための一時チャート作成
    Dim cht As ChartObject
    Set cht = ActiveSheet.ChartObjects.Add(Left:=0, Top:=0, Width:=100, Height:=100)
    
    ' 画像をチャートに追加
    cht.Chart.Paste
    
    ' チャートを画像としてエクスポート
    cht.Chart.Export filePath
    
    ' 一時チャートを削除
    cht.Delete
End Sub
