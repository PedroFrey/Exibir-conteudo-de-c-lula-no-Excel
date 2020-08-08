'Prints out formula as text'Maja Sliwinski developed this for Simon Benninga
Function GetFormula(Rng As Range) As String
    Application.Volatile True
    GetFormula = "<-- " & Application.Text(Rng.FormulaLocal, "")
End Function
