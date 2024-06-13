Sub ResizeImagesWidthForVBA()
    Dim iShape As InlineShape
    Dim newWidthMM As Single
    Dim newWidthPoints As Single
    
    ' 新しい幅を設定します（単位：ミリメートル）
    newWidthMM = 50 ' 幅をミリメートルで指定
    
    ' ミリメートルをポイントに変換（1ポイント = 0.352777ミリメートル）
    newWidthPoints = newWidthMM / 0.352777
    
    ' ドキュメント内のすべてのインライン画像をループします
    On Error Resume Next ' エラーが発生しても次の画像に進むように設定
    For Each iShape In ActiveDocument.InlineShapes
        With iShape
            .LockAspectRatio = msoTrue
            .Width = newWidthPoints
            If Err.Number <> 0 Then
                MsgBox "エラーが発生しました。画像のサイズ変更に失敗しました。" & vbCrLf & "エラー番号: " & Err.Number & vbCrLf & "エラーメッセージ: " & Err.Description, vbExclamation
                Err.Clear ' エラーをクリアして次の画像に進む
            End If
        End With
    Next iShape
    On Error GoTo 0 ' エラーハンドリングを元に戻す
    
    ' 処理完了を通知
    MsgBox "すべての画像のサイズ変更が完了しました。", vbInformation
End Sub
