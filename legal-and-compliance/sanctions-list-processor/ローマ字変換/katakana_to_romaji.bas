Sub KatakanaToRomaji()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim katakanaText As String
    Dim romajiText As String
    
    ' アクティブシートを設定
    Set ws = ActiveSheet
    
    ' C列の最終行を取得
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' C列の各セルを処理
    For i = 1 To lastRow
        katakanaText = ws.Cells(i, 3).Value
        If katakanaText <> "" Then
            romajiText = ConvertToRomaji(katakanaText)
            ws.Cells(i, 3).Value = romajiText
        End If
    Next i
    
    MsgBox "変換が完了しました。", vbInformation
End Sub

Function ConvertToRomaji(katakana As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim nextChar As String
    
    result = ""
    i = 1
    
    Do While i <= Len(katakana)
        char = Mid(katakana, i, 1)
        
        ' 次の文字を確認（小文字の処理用）
        If i < Len(katakana) Then
            nextChar = Mid(katakana, i + 1, 1)
        Else
            nextChar = ""
        End If
        
        ' 2文字の組み合わせを先にチェック
        If i < Len(katakana) Then
            Select Case char & nextChar
                ' 拗音
                Case "キャ": result = result & "kya": i = i + 1
                Case "キュ": result = result & "kyu": i = i + 1
                Case "キョ": result = result & "kyo": i = i + 1
                Case "シャ": result = result & "sha": i = i + 1
                Case "シュ": result = result & "shu": i = i + 1
                Case "ショ": result = result & "sho": i = i + 1
                Case "チャ": result = result & "cha": i = i + 1
                Case "チュ": result = result & "chu": i = i + 1
                Case "チョ": result = result & "cho": i = i + 1
                Case "ニャ": result = result & "nya": i = i + 1
                Case "ニュ": result = result & "nyu": i = i + 1
                Case "ニョ": result = result & "nyo": i = i + 1
                Case "ヒャ": result = result & "hya": i = i + 1
                Case "ヒュ": result = result & "hyu": i = i + 1
                Case "ヒョ": result = result & "hyo": i = i + 1
                Case "ミャ": result = result & "mya": i = i + 1
                Case "ミュ": result = result & "myu": i = i + 1
                Case "ミョ": result = result & "myo": i = i + 1
                Case "リャ": result = result & "rya": i = i + 1
                Case "リュ": result = result & "ryu": i = i + 1
                Case "リョ": result = result & "ryo": i = i + 1
                Case "ギャ": result = result & "gya": i = i + 1
                Case "ギュ": result = result & "gyu": i = i + 1
                Case "ギョ": result = result & "gyo": i = i + 1
                Case "ジャ": result = result & "ja": i = i + 1
                Case "ジュ": result = result & "ju": i = i + 1
                Case "ジョ": result = result & "jo": i = i + 1
                Case "ビャ": result = result & "bya": i = i + 1
                Case "ビュ": result = result & "byu": i = i + 1
                Case "ビョ": result = result & "byo": i = i + 1
                Case "ピャ": result = result & "pya": i = i + 1
                Case "ピュ": result = result & "pyu": i = i + 1
                Case "ピョ": result = result & "pyo": i = i + 1
                Case "ヂャ": result = result & "dya": i = i + 1
                Case "ヂュ": result = result & "dyu": i = i + 1
                Case "ヂョ": result = result & "dyo": i = i + 1
                
                ' その他の特殊な組み合わせ
                Case "ファ": result = result & "fa": i = i + 1
                Case "フィ": result = result & "fi": i = i + 1
                Case "フェ": result = result & "fe": i = i + 1
                Case "フォ": result = result & "fo": i = i + 1
                Case "ヴァ": result = result & "va": i = i + 1
                Case "ヴィ": result = result & "vi": i = i + 1
                Case "ヴェ": result = result & "ve": i = i + 1
                Case "ヴォ": result = result & "vo": i = i + 1
                Case "ウィ": result = result & "wi": i = i + 1
                Case "ウェ": result = result & "we": i = i + 1
                Case "ウォ": result = result & "wo": i = i + 1
                Case "ティ": result = result & "ti": i = i + 1
                Case "ディ": result = result & "di": i = i + 1
                Case "トゥ": result = result & "tu": i = i + 1
                Case "ドゥ": result = result & "du": i = i + 1
                
                Case Else
                    ' 1文字の処理へ
                    GoTo SingleChar
            End Select
        Else
            GoTo SingleChar
        End If
        
        GoTo ContinueLoop
        
SingleChar:
        ' 1文字の変換
        Select Case char
            ' 基本の五十音
            Case "ア": result = result & "a"
            Case "イ": result = result & "i"
            Case "ウ": result = result & "u"
            Case "エ": result = result & "e"
            Case "オ": result = result & "o"
            Case "カ": result = result & "ka"
            Case "キ": result = result & "ki"
            Case "ク": result = result & "ku"
            Case "ケ": result = result & "ke"
            Case "コ": result = result & "ko"
            Case "サ": result = result & "sa"
            Case "シ": result = result & "shi"
            Case "ス": result = result & "su"
            Case "セ": result = result & "se"
            Case "ソ": result = result & "so"
            Case "タ": result = result & "ta"
            Case "チ": result = result & "chi"
            Case "ツ": result = result & "tsu"
            Case "テ": result = result & "te"
            Case "ト": result = result & "to"
            Case "ナ": result = result & "na"
            Case "ニ": result = result & "ni"
            Case "ヌ": result = result & "nu"
            Case "ネ": result = result & "ne"
            Case "ノ": result = result & "no"
            Case "ハ": result = result & "ha"
            Case "ヒ": result = result & "hi"
            Case "フ": result = result & "fu"
            Case "ヘ": result = result & "he"
            Case "ホ": result = result & "ho"
            Case "マ": result = result & "ma"
            Case "ミ": result = result & "mi"
            Case "ム": result = result & "mu"
            Case "メ": result = result & "me"
            Case "モ": result = result & "mo"
            Case "ヤ": result = result & "ya"
            Case "ユ": result = result & "yu"
            Case "ヨ": result = result & "yo"
            Case "ラ": result = result & "ra"
            Case "リ": result = result & "ri"
            Case "ル": result = result & "ru"
            Case "レ": result = result & "re"
            Case "ロ": result = result & "ro"
            Case "ワ": result = result & "wa"
            Case "ヲ": result = result & "wo"
            Case "ン": result = result & "n"
            
            ' 濁音
            Case "ガ": result = result & "ga"
            Case "ギ": result = result & "gi"
            Case "グ": result = result & "gu"
            Case "ゲ": result = result & "ge"
            Case "ゴ": result = result & "go"
            Case "ザ": result = result & "za"
            Case "ジ": result = result & "ji"
            Case "ズ": result = result & "zu"
            Case "ゼ": result = result & "ze"
            Case "ゾ": result = result & "zo"
            Case "ダ": result = result & "da"
            Case "ヂ": result = result & "ji"
            Case "ヅ": result = result & "zu"
            Case "デ": result = result & "de"
            Case "ド": result = result & "do"
            Case "バ": result = result & "ba"
            Case "ビ": result = result & "bi"
            Case "ブ": result = result & "bu"
            Case "ベ": result = result & "be"
            Case "ボ": result = result & "bo"
            
            ' 半濁音
            Case "パ": result = result & "pa"
            Case "ピ": result = result & "pi"
            Case "プ": result = result & "pu"
            Case "ペ": result = result & "pe"
            Case "ポ": result = result & "po"
            
            ' 小文字
            Case "ァ": result = result & "a"
            Case "ィ": result = result & "i"
            Case "ゥ": result = result & "u"
            Case "ェ": result = result & "e"
            Case "ォ": result = result & "o"
            Case "ッ"
                ' 促音の処理
                If i < Len(katakana) Then
                    Dim nextRomaji As String
                    nextRomaji = ConvertSingleChar(Mid(katakana, i + 1, 1))
                    If Left(nextRomaji, 1) <> "" And nextRomaji <> Mid(katakana, i + 1, 1) Then
                        result = result & Left(nextRomaji, 1)
                    End If
                End If
            Case "ー": result = result & "-"
            Case "ヴ": result = result & "vu"
            
            ' その他（変換できない文字はそのまま）
            Case Else: result = result & char
        End Select
        
ContinueLoop:
        i = i + 1
    Loop
    
    ConvertToRomaji = result
End Function

Function ConvertSingleChar(char As String) As String
    ' 促音処理用の単一文字変換関数
    Select Case char
        Case "カ": ConvertSingleChar = "ka"
        Case "キ": ConvertSingleChar = "ki"
        Case "ク": ConvertSingleChar = "ku"
        Case "ケ": ConvertSingleChar = "ke"
        Case "コ": ConvertSingleChar = "ko"
        Case "サ": ConvertSingleChar = "sa"
        Case "シ": ConvertSingleChar = "shi"
        Case "ス": ConvertSingleChar = "su"
        Case "セ": ConvertSingleChar = "se"
        Case "ソ": ConvertSingleChar = "so"
        Case "タ": ConvertSingleChar = "ta"
        Case "チ": ConvertSingleChar = "chi"
        Case "ツ": ConvertSingleChar = "tsu"
        Case "テ": ConvertSingleChar = "te"
        Case "ト": ConvertSingleChar = "to"
        Case "ハ": ConvertSingleChar = "ha"
        Case "ヒ": ConvertSingleChar = "hi"
        Case "フ": ConvertSingleChar = "fu"
        Case "ヘ": ConvertSingleChar = "he"
        Case "ホ": ConvertSingleChar = "ho"
        Case "パ": ConvertSingleChar = "pa"
        Case "ピ": ConvertSingleChar = "pi"
        Case "プ": ConvertSingleChar = "pu"
        Case "ペ": ConvertSingleChar = "pe"
        Case "ポ": ConvertSingleChar = "po"
        Case "ガ": ConvertSingleChar = "ga"
        Case "ギ": ConvertSingleChar = "gi"
        Case "グ": ConvertSingleChar = "gu"
        Case "ゲ": ConvertSingleChar = "ge"
        Case "ゴ": ConvertSingleChar = "go"
        Case "ザ": ConvertSingleChar = "za"
        Case "ジ": ConvertSingleChar = "ji"
        Case "ズ": ConvertSingleChar = "zu"
        Case "ゼ": ConvertSingleChar = "ze"
        Case "ゾ": ConvertSingleChar = "zo"
        Case "ダ": ConvertSingleChar = "da"
        Case "ヂ": ConvertSingleChar = "ji"
        Case "ヅ": ConvertSingleChar = "zu"
        Case "デ": ConvertSingleChar = "de"
        Case "ド": ConvertSingleChar = "do"
        Case "バ": ConvertSingleChar = "ba"
        Case "ビ": ConvertSingleChar = "bi"
        Case "ブ": ConvertSingleChar = "bu"
        Case "ベ": ConvertSingleChar = "be"
        Case "ボ": ConvertSingleChar = "bo"
        Case Else: ConvertSingleChar = char
    End Select
End Function