'First Function

Function PVDA(r, c, m, g, i, n)

'PVDA computes the value of delayed perpetuity_or delayed annuity by choosing the type in the function

'r is interest rate
'c is the amount of payment
'm is the length of delay for initial payment
'g is growth rate for periodic payments
'i is type of valuation (=1 when type is delayed perpetuity, i = 2 when type is delayed annuity)
'n = number of annuity payments

'Report missing inputs:

    If IsMissing(r) Then
        msg = "Please write the interest rate appropriately (as %)"
        MsgBox msg
    ElseIf IsMissing(c) Then
        msg = "Please write the amount of payment appropriately"
        MsgBox msg
    ElseIf IsMissing(m) Then
        msg = "Please write the length of delay for initial payment appropriately"
        MsgBox msg
    ElseIf IsMissing(g) Then
        msg = "Please write the growth rate appropriately or 0% if there is no growth rate"
        MsgBox msg
    ElseIf IsMissing(i) Then
        msg = "Please write type of valuation appropriately (1 if delayed perpetuity, 2 if delayed annuity"
        MsgBox msg
    ElseIf i = 2 Then
        If IsMissing(n) Then
        msg = "Please write number of annuity payment appropriately"
        MsgBox msg
        End If
        
  End If

    
    If i = 1 Then 'perpetuity type
      PVDA = 1 / (1 + r) ^ m * c / (r - g)
    ElseIf i = 2 Then 'annuity type
      PVDA = 1 / (1 + r) ^ m * c / (r - g) * (1 - ((1 + g) / (1 + r)) ^ n)
    Else
      msg1 = "Please write type of valuation correctly (1 if delayed perpetuity, 2 if delayed annuity)"
      MsgBox msg1
    End If
    
End Function


'Second Function:


Function Return_risk(w1, w2, w3, mean1, mean2, mean3, Optional std1, Optional std2, Optional std3, Optional corr12, Optional corr13, Optional corr23)

'Return_risk function finds expected return of three stocks if you do not give standard values of stocks and also finds standard deviation of the three stocks with given weights
 
If IsMissing(std1) Then

    expected_return = w1 * mean1 + w2 * mean2 + w3 * mean3

    result = expected_return

Else

    Var = w1 * w1 * std1 * std1 + w2 * w2 * std2 * std2 + w3 * w3 * std3 * std3 + 2 * w1 * w2 * std1 * std2 * corr12 + 2 * w1 * w3 * std1 * std3 * corr13 + 2 * w2 * w3 * std2 * std3 * corr23

    stdofreturn = Sqr(Var)

    result = stdofreturn

End If

    Return_risk = result

End Function


'Third Function:


Function VarCovar(rng As Range) As Variant

'VarCovar function finds variance-covariance matrix of returns of stocks

    Dim i As Integer
    Dim j As Integer
    Dim numcols As Integer
    numcols = rng.Columns.Count
    numrows = rng.Rows.Count
    Dim matrix() As Double
    ReDim matrix(numcols - 1, numcols - 1)

    For i = 1 To numcols
        For j = 1 To numcols
            matrix(i - 1, j - 1) = Application.WorksheetFunction.Covariance_S(rng.Columns(i), rng.Columns(j))
        Next j
    Next i
    VarCovar = matrix
    
End Function

