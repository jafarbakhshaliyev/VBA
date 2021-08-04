'First Function:

Function dOne(Stock, Exercise, Time, Interest, sigma)

'dOne function finds d1 value in Black-Scholes formula for finding call value C

dOne = (Log(Stock / Exercise) + Interest * Time) / (sigma _
    * Sqr(Time)) + 0.5 * sigma * Sqr(Time)
End Function

'Second Function:

Function BSCall(Stock, Exercise, Time, Interest, sigma)

'Finds Black-Scholes value of call option

BSCall = Stock * Application.NormSDist(dOne(Stock, _
    Exercise, Time, Interest, sigma)) - Exercise * _
    Exp(-Time * Interest) * Application.NormSDist _
    (dOne(Stock, Exercise, Time, Interest, sigma) _
    - sigma * Sqr(Time))
    
End Function

'Third Function:

Function VanillaCall(S0, Exercise, Mean, sigma, _
Interest, Time, Divisions, Runs)

'VanillaCall finds value of vanilla call option

    deltat = 1 / Divisions
    interestdelta = Exp(Interest * deltat)
    
    up = Exp(Mean * deltat + _
    sigma * Sqr(deltat))
    down = Exp(Mean * deltat - _
    sigma * Sqr(deltat))
    
    pathlength = Int(Time / deltat)
    
'Risk-neutral probabilities
piup = (interestdelta - down) / _
(up - down)
pidown = 1 - piup
         
Temp = 0

For Index = 1 To Runs
    Upcounter = 0
    'Generate terminal price
    For j = 1 To pathlength
    If Rnd > pidown Then Upcounter = _
    Upcounter + 1
    Next j
    callvalue = Application.Max(S0 * _
    (up ^ Upcounter) * (down ^ (pathlength - _
    Upcounter)) - Exercise, 0) _
    / (interestdelta ^ pathlength)
    Temp = Temp + callvalue
Next Index

VanillaCall = Temp / Runs

End Function


