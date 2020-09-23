Attribute VB_Name = "modRandomGen"
Private Z() As Long
Private Max
Private Min
Private Up As Long
Private Low As Long
Private n As Long
Private x As Long

Public Function Rand(ByVal Low As Long, ByVal High As Long) As Long
    Randomize
    Rand = Int((High - Low + 1) * Rnd) + Low
End Function

Public Function kRandomize(LowerBound As Long, UpperBound As Long)
    Dim i As Long
    ReDim Z(0 To ((UpperBound - LowerBound) - 1))

    x = 0
    Up = UpperBound
    Low = LowerBound
    
    Max = UpperBound
    Min = LowerBound
    Randomize
    
    For n = 0 To (UpperBound - LowerBound) - 1
        Z(n) = Shuffle(Min, Max)
        RemoveDoubles
    Next n
        
End Function

Public Function kRnd() As Long
    kRnd = Z(x)
    If x = ((Up - Low) - 1) Then
        kRandomize Low, Up
    Else
        x = x + 1
    End If
End Function

Public Sub RemoveDoubles()
Dim i As Long

For i = 0 To n
    If n = 0 Then Exit For
    If Z(n) = Z(i) And n <> i Then
        Z(n) = "0"
        n = n - 1
    End If
Next i

End Sub

Static Function Shuffle(ByVal Lower As Long, ByVal Upper As Long) As Long
Static PrimeFactor(10) As Long
Static A As Long, c As Long, B As Long
Static s As Long, n As Long
Dim i As Long, J As Long, k As Long
Dim m As Long

    If (n <> Upper - Lower + 1) Then
        n = Upper - Lower + 1
        i = 0
        n1 = n
        k = 2
    
    
        Do While k <= n1
    
    
            If (n1 Mod k = 0) Then
    
    
                If (i = 0 Or PrimeFactor(i) <> k) Then
                    i = i + 1
                    PrimeFactor(i) = k
                End If
                n1 = n1 / k
            Else
                k = k + 1
            End If
        Loop
        B = 1
    
    
        For J = 1 To i
            B = B * PrimeFactor(J)
        Next J
        If n Mod 4 = 0 Then B = B * 2
        A = B + 1
        c = Int(n * 0.66)
        t = True
    
    
        Do While t
            t = False
    
    
            For J = 1 To i
                If ((c Mod PrimeFactor(J) = 0) Or (c Mod A = 0)) Then t = True
            Next J
            If t Then c = c - 1
        Loop
        Randomize
        s = Rnd(n)
    End If
    s = (A * s + c) Mod n
    Shuffle = s + Lower
End Function

