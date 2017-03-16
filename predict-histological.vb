Option Explicit

Sub caculate()

Dim age, size, nodes, grade, erstat, detection, chemoGen, her2, ki67 As Double
Dim i As Long
Dim myarray As Range
Dim year5_results, year10_results As Range

    Set myarray = Range(Range("A2"), Range("I2").End(xlDown))
    Set year5_results = Range(Cells(2, 10), Cells(myarray.Rows.Count + 1, 15))
    Set year10_results = Range(Cells(2, 16), Cells(myarray.Rows.Count + 1, 21))
    
    i = 2
    
    For i = 2 To myarray.Rows.Count + 1
        age = Range("A" & i).Value
        detection = Range("B" & i).Value
        size = Range("C" & i).Value
        grade = Range("D" & i).Value
        nodes = Range("E" & i).Value
        erstat = Range("F" & i).Value
        her2 = Range("G" & i).Value
        ki67 = Range("H" & i).Value
        chemoGen = Range("I" & i).Value
               
        If erstat = 1 Or erstat = 0 Then
            year5_results.Rows(i - 1).Value = predict_v2_0(age, detection, size, grade, nodes, erstat, her2, ki67, chemoGen, 5)
            year10_results.Rows(i - 1).Value = predict_v2_0(age, detection, size, grade, nodes, erstat, her2, ki67, chemoGen, 10)
        End If
    
    Next i
     
End Sub
   
' Arguments age, size and nodes are entered as values; the others as lookups
' # This is how the model assigns some input parameters (or ranges) into variables
' # i.e. parameter (or ranges) -> web form setting -> Predict model variable setting
' # Tumour Grade (1,2,3,unknown) -> (1,2,3,9) -> (1.0,2.0,3.0,2.13)
' # ER Status (-ve,+ve) -> (0,1) -> (0,1) n.b. unknown not allowed
' # Detection (Clinical,Screening,Other) -> (0,1,2) -> (0.0,1.0,0.204)
' # Chemo (1st,2nd,3rd) -> (1,2,3) -> (1,2,3)
' # HER2 Status (-ve,+ve,unknown) -> (0,1,9) -> (0,1,9) n.b. these are now changed in web code (were (1,2,0))
' # KI67 Status (-ve,+ve,unknown) -> (0,1,9) -> (0,1,9) n.b. these are now changed in web code (were (1,2,0))



Function predict_v2_0(age, detection, size, grade, nodes, erstat, her2, ki67, chemoGen, rtime As Double) As Variant


'prevent high hazard for young patients - age of 24.54 caps Hazard ratio at 4.008


If age < 25 Then
    age = 25
End If
    
Dim grade_a As Double
    grade_a = 0
    If grade = 2 Or grade = 3 Then
        grade_a = 1
    End If
    
'console.log("check grade,detection: ",grade,detection);
'n.b. default of 0 for her2_rh and ki67_rh remains when inputs are set as undefined
       
       
'her_rh
Dim her2_rh As Double
    her2_rh = 0
    
    If her2 = 1 And erstat = 1 Then
        her2_rh = 0.2413
    ElseIf her2 = 0 And erstat = 1 Then
        her2_rh = -0.0762
    ElseIf her2 = 1 And erstat = 0 Then
        her2_rh = 0.2413
    ElseIf her2 = 0 And erstat = 0 Then
        her2_rh = -0.0762
    End If
    
    
'ki67_rh
Dim ki67_rh As Double
    ki67_rh = 0
    
'If ki67 = 1 And erstat = 1 Then
       ' ki67_rh = 0.14904
   ' Else: ki67 = 0 And erstat = 1
       ' ki67_rh = -0.11333
   ' End If
    
'c
'No chemo - use initial value of c i.e. 0
'First chemoGen - not used here for completeness

Dim c As Double
    c = 0
    If chemoGen = 1 Then
        If age < 50 And erstat = 0 Then
            c = -0.3567
        ElseIf age >= 50 And erstat = 0 Then
            c = -0.2485
        ElseIf age >= 60 And erstat = 0 Then
            c = -0.1278
        ElseIf age < 50 And erstat = 1 Then
            c = -0.3567
        ElseIf age >= 50 And erstat = 1 Then
            c = -0.1744
        ElseIf age >= 60 And erstat = 1 Then
            c = -0.0834
        End If
    ElseIf chemoGen = 2 Then
        c = -0.248
    ElseIf chemoGen = 3 Then
        c = -0.446
    End If


Dim surv_oth_time() As Variant
ReDim surv_oth_time(0 To rtime)
surv_oth_time(0) = 0
Dim surv_oth_year() As Variant
ReDim surv_oth_year(0 To rtime)
surv_oth_year(0) = 0
Dim bs() As Variant
ReDim bs(0 To rtime)
bs(0) = 0

'The number of columns is not important, as it is not required to specify the size of an array before using it

Dim ftime As Long
    ftime = Round(rtime)
Dim mort_rate_cum_rx() As Double
Dim surv_br_time_rx() As Double
Dim pr_all_time_rx() As Double
Dim pr_oth_time_rx() As Double
Dim pr_br_time_rx() As Double
Dim pr_dfs_time_rx() As Double
ReDim mort_rate_cum_rx(0 To 9, 0 To ftime + 1)
ReDim surv_br_time_rx(0 To 9, 0 To ftime + 1)
ReDim pr_all_time_rx(0 To 9, 0 To ftime + 1)
ReDim pr_oth_time_rx(0 To 9, 0 To ftime + 1)
ReDim pr_br_time_rx(0 To 9, 0 To ftime + 1)
ReDim pr_dfs_time_rx(0 To 9, 0 To ftime + 1)


Dim benefit_h, benefit_h10, benefit_c, benefit_t, benefit_hc, benefit_h10c, benefit_hct, benefit_h10ct As Double


Dim time As Long
time = 1
For time = 1 To ftime


' Calculate the breast cancer mortality prognostic index (pi)
' Generate baseline survival
Dim pi As Double

pi = 0

If erstat = 1 Then
    pi = 34.53642 * (Application.WorksheetFunction.Power(age / 10, -2) - 0.0287449295) _
        - 34.20342 * (Application.WorksheetFunction.Power(age / 10, -2) * Log(age / 10) - 0.0510121013) _
        + 0.7530729 * (Log(size / 100) + 1.545233938) _
        + 0.7060723 * (Log((nodes + 1) / 10) + 1.387566896) _
        + 0.746655 * grade _
        - 0.22763366 * detection _
        + her2_rh _
        + ki67_rh

  bs(time) = Exp(0.7424402 - 7.527762 * (Application.WorksheetFunction.Power(1 / time, 0.5)) - 1.812513 * Application.WorksheetFunction.Power(1 / time, 0.5) * Log(time))
    
ElseIf erstat = 0 Then
    pi = 0.0089827 * (age - 56.3254902) _
       + 2.093446 * (Application.WorksheetFunction.Power(size / 100#, 0.5) - 0.5090456276) _
       + 0.6260541 * (Log((nodes + 1) / 10#) + 1.086916249) _
       + 1.129091 * grade_a _
       + her2_rh + ki67_rh
  bs(time) = Exp(-1.156036 + 0.4707332 / Application.WorksheetFunction.Power(time, 2#) - 3.51355 / time)
End If

'Generate therapy reduction coefficients
    
    
Dim h As Double
Dim t As Double

If erstat = 1 Then
    h = -0.3857
End If

Dim h10 As Double
h10 = h

If erstat = 1 And time > 10 Then
    h10 = h - 0.3425
End If

Dim hc, hct, h10c, h10ct As Double
hc = h + c
hct = h + c + t
h10c = h10 + c
h10ct = h10 + c + t

Dim types() As Variant
    types = Array(0, h, h10, c, t, hc, h10c, hct, h10ct)

'Generate cumulative survival non-breast mortality
Dim bs_oth As Double
bs_oth = Math.Exp(-6.052919 + 1.079863 * Log(time) + 0.3255321 * Application.WorksheetFunction.Power(time, 0.5))
surv_oth_time(time) = Math.Exp(-Math.Exp(0.0698252 * (Application.WorksheetFunction.Power(age / 10#, 2#) - 34.23391957)) * bs_oth)


'Generate annual survival from cumulative survival
If time = 1 Then
    surv_oth_year(time) = 1 - surv_oth_time(time)
ElseIf time > 1 Then
    surv_oth_year(time) = surv_oth_time(time - 1) - surv_oth_time(time)
End If


Dim mort_rate_rx, surv_br_year_rx, pr_all_year_rx, proportion_br_rx, pr_oth_year_rx, pr_br_year_rx As Double
Dim pr_oth_time_0 As Double
Dim pr_br_time_0 As Double
Dim pr_dfs_time_0 As Double

Dim i As Double
i = 0

For i = 0 To 8

        Dim rx As Double
            rx = types(i)
        
        'Generate the breast cancer specific survival
        
        If time = 1 Then
            mort_rate_rx = bs(time) * Exp(pi + rx)
            mort_rate_cum_rx(i, time) = mort_rate_rx
            surv_br_year_rx = 1 - surv_br_time_rx(i, time)
        ElseIf time > 1 Then
            mort_rate_rx = (bs(time) - bs(time - 1)) * Math.Exp(pi + rx)
            mort_rate_cum_rx(i, time) = mort_rate_rx + mort_rate_cum_rx(i, time - 1)
            surv_br_time_rx(i, time) = Math.Exp(-mort_rate_cum_rx(i, time))
            surv_br_year_rx = surv_br_time_rx(i, time - 1) - surv_br_time_rx(i, time)
        End If
        'All cause mortality
        
        pr_all_time_rx(i, time) = 1 - surv_oth_time(time) * surv_br_time_rx(i, time)
        'Cumulative all cause mortality
        If time = 1 Then
            pr_all_year_rx = pr_all_time_rx(i, time)
        End If
        'Number deaths in year 1
        If time > 1 Then
            pr_all_year_rx = pr_all_time_rx(i, time) - pr_all_time_rx(i, time - 1)
        End If
        'Proportion of all cause mortality
        
        proportion_br_rx = (surv_br_year_rx) / (surv_oth_year(time) + surv_br_year_rx)
        pr_oth_year_rx = (1 - proportion_br_rx) * pr_all_year_rx
        If time = 1 Then
            pr_oth_time_rx(i, time) = pr_oth_year_rx
        End If
            
        If time > 1 Then
            pr_oth_time_rx(i, time) = pr_oth_year_rx + pr_oth_time_rx(i, time - 1)
        End If
        
        'Breast mortality and recurrence as competing risk
        
        pr_br_year_rx = proportion_br_rx * pr_all_year_rx
        If time = 1 Then
            pr_br_time_rx(i, time) = pr_br_year_rx
        End If
        
        If time > 1 Then
            pr_br_time_rx(i, time) = pr_br_year_rx + pr_br_time_rx(i, time - 1)
        End If
        
        pr_dfs_time_rx(i, time) = 1 - Math.Exp(Log(1 - pr_br_time_rx(i, time)) * 1.3)
        
        'assign results at final time i.e. time required for results
        If time = ftime Then
            If i = 0 Then
                pr_oth_time_0 = pr_oth_time_rx(i, time)
                pr_br_time_0 = pr_br_time_rx(i, time)
                pr_dfs_time_0 = pr_dfs_time_rx(i, time)
            'Benefits of treatment
            ElseIf i = 1 Then
                benefit_h = pr_br_time_0 - pr_br_time_rx(i, time)
            ElseIf i = 2 Then
                benefit_h10 = pr_br_time_0 - pr_br_time_rx(i, time)
            ElseIf i = 3 Then
                benefit_c = pr_br_time_0 - pr_br_time_rx(i, time)
            ElseIf i = 4 Then
                benefit_t = pr_br_time_0 - pr_br_time_rx(i, time)
            ElseIf i = 5 Then
                benefit_hc = pr_br_time_0 - pr_br_time_rx(i, time)
            ElseIf i = 6 Then
                benefit_h10c = pr_br_time_0 - pr_br_time_rx(i, time)
            ElseIf i = 7 Then
                benefit_hct = pr_br_time_0 - pr_br_time_rx(i, time)
            Else: i = 8
                benefit_h10ct = pr_br_time_0 - pr_br_time_rx(i, time)
            End If
        End If
                
                    
        'gen benefit_rx = pr_br_time_0(time) - pr_br_time_rx(time)
        'gen benefit_dfs_rx = pr_dfs_time_0(time) - pr_dfs_time_rx(time)
        
        
        
            'end of types loop h,c etc
Next i


'end of time loop
Next time


'Additive Benefits - these are hierarchical


Dim bcSpecSur, cumOverallSurOL, cumOverallSurHormo, cumOverallSurChemo, cumOverallSurCandH, cumOverallSurCHT As Double
    bcSpecSur = 1 - pr_br_time_0
'outputs required for predict bar chart display
    cumOverallSurOL = 1# - pr_oth_time_0 - pr_br_time_0
    cumOverallSurHormo = benefit_h
    cumOverallSurChemo = benefit_c
    cumOverallSurCandH = benefit_hc
    cumOverallSurCHT = benefit_hct

'n.b. this is original predict (V1) return line
'return (bcSpecSur, cumOverallSurOL, cumOverallSurChemo, cumOverallSurHormo, cumOverallSurCandH, cumOverallSurCHT,
'          pySurv10OL, pySurv10Chemo, pySurv10Hormo, pySurv10CandH, pySurv10CHT);
'n.b.this bcSpecSur Is Not displayed

predict_v2_0 = Array(bcSpecSur, cumOverallSurOL, cumOverallSurHormo, cumOverallSurChemo, cumOverallSurCandH, cumOverallSurCHT)
    
End Function
