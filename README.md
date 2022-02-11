# Useful Excel
Useful Excel Formulas &amp; VBA


<h2>Range to list</h2>

Simple formula to convert a range into a list in one column

<code>=INDEX(Data,1+INT((ROW(A1)-1)/COLUMNS(Data)),MOD(ROW(A1)-1+COLUMNS(Data),COLUMNS(Data))+1)</code>


<h2>Multi Replace</h2>

VBA Macro to make a series of find and replaces from lists

- MatchCase can be changed to True to only change when case matches
- xlPart can be changed to xlWhole to only replace an exact cell

Sub MultiReplace()
On Error GoTo errorcatch
Dim arrRules() As Variant

    strSheet = InputBox("Enter sheet name where your replace rules are", _
        "Sheet name", "Sheet1")
    strRules = InputBox("Enter address of replaces rules." & vbNewLine & _
        "But only the first column!", "Address", "A1:A100")

    Set rngCol1 = Sheets(strSheet).Range(strRules)
    Set rngCol2 = rngCol1.Offset(0, 1)
    arrRules = Application.Union(rngCol1, rngCol2)

    For i = 1 To UBound(arrRules)
        Selection.Replace What:=arrRules(i, 1), Replacement:=arrRules(i, 2), _
            LookAt:=xlPart, MatchCase:=False
    Next i

errorcatch:
End Sub


