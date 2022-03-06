Dim WithEvents RC As SpInProcRecoContext
Dim Recognizer As SpInprocRecognizer
Dim myGrammar As ISpeechRecoGrammar

Private Sub Form_Load()
    On Error GoTo EH

    Set RC = New SpInProcRecoContext
    Set Recognizer = RC.Recognizer

    Set myGrammar = RC.CreateGrammar
    myGrammar.DictationSetState SGDSActive

    Dim Category As SpObjectTokenCategory
    Set Category = New SpObjectTokenCategory
    Category.SetId SpeechCategoryAudioIn

    Dim Token As SpObjectToken
    Set Token = New SpObjectToken
    Token.SetId Category.Default()
    Set Recognizer.AudioInput = Token

EH:
    If Err.Number Then ShowErrMsg
End Sub

Private Sub RC_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)
    On Error GoTo EH

    Label1.Caption = Result.PhraseInfo.GetText

EH:
    If Err.Number Then ShowErrMsg
End Sub


Private Sub ShowErrMsg()

    ' Declare identifiers:
    Const NL = vbNewLine
    Dim T As String

    T = "Desc: " & Err.Description & NL
    T = T & "Err #: " & Err.Number
    MsgBox T, vbExclamation, "Run-Time Error"
    End

End Sub