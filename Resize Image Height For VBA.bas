Sub ResizeImageHeightForVBA()
    Dim iShape As InlineShape
    Dim newHeightMM As Single
    Dim newHeightPoints As Single
    
    ' 新しい高さを設定します（単位：ミリメートル）
    newHeightMM = 50 ' 高さをミリメートルで指定
    
    ' ミリメートルをポイントに変換（1ポイント = 0.352777ミリメートル）
    newHeightPoints = newHeightMM / 0.352777
    
    ' ドキュメント内のすべてのインライン画像をループします
    On Error Resume Next ' エラーが発生しても次の画像に進むように設定
    For Each iShape In ActiveDocument.InlineShapes
        With iShape
            .LockAspectRatio = msoFalse ' アスペクト比のロックを解除
            .Height = newHeightPoints
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
