Attribute VB_Name = "std_Chemistry"


Option Explicit

Public Type std_Chemistry_Element
    AtomicNumber As Long
    Name         As String
    Short        As String
    MolarMass    As Double
End Type

Public Enum std_Chemistry_Element_Property
    All          = 0
    AtomicNumber = 1
    Name         = 2
    Short        = 3
    MolarMass    = 4
End Enum

Private PSE(118) As std_Chemistry_Element


Public Function std_Chemistry_Intepret(Text As String, ReturnWhat As std_Chemistry_Element_Property) As Variant()

    Dim StartPos              As Long   : StartPos = 1
    Dim EndPos                As Long   : EndPos = Len(Text)
    Dim ReturnArray()         As Variant: ReDim ReturnArray(1, 0)
    Dim Depth                 As Long   : Depth = -1
    Dim MaxDepth              As Long   : MaxDepth = Depth
    Dim Temp()                As Variant: ReDim Temp(1, 100, 0)
    Dim CurrentText           As String : CurrentText = ""
    Dim Found                 As Boolean: Found = False
    Dim Success               As Boolean: Success = False
    Dim ElementIndex          As Long
    Dim i As Long, j As Long, k As Long


    Call Initialize()
    For i = Len(Text) To 0 Step -1
        CurrentText = Mid(Text, StartPos, EndPos)
        Do Until Len(CurrentText) = 0
            CurrentText = Mid(Text, StartPos, EndPos)
            Select Case True
                Case IsNumeric(CurrentText)
                    If Mid(Text, StartPos - 1, 1) = ")" Then
                        For k = 0 To Ubound(Temp, 3) - 1
                            Temp(1, Depth + 1, k) = CDbl(CurrentText) * CDbl(Temp(1, Depth + 1, k))
                        Next k
                    Else
                        If Depth >= 0 Then
                            Temp(1, Depth, Ubound(Temp, 3) - 1) = CDbl(CurrentText)
                        Else
                            ReturnArray(1, UBound(ReturnArray, 2) - 1) = CDbl(CurrentText)
                        End If
                    End If
                    Success = True
                Case CurrentText = "(" : Success = True: Depth = Depth + 1: MaxDepth = MaxDepth + 1
                Case CurrentText = ")" : Success = True: Depth = Depth - 1
                Case CurrentText = ""  : Success = True
                Case Else
                    ElementIndex = std_Chemistry_SearchPSE(CurrentText, std_Chemistry_Element_Property.Short)
                    If ElementIndex <> 0 Then
                        If Depth >= 0 Then
                            Temp(0, Depth, Ubound(Temp, 3)) = std_Chemistry_GetPSE(ElementIndex, ReturnWhat)
                            Temp(1, Depth, Ubound(Temp, 3)) = 1
                            ReDim Preserve Temp(1, 100, Ubound(Temp, 3) + 1)
                        Else
                            ReturnArray(0, UBound(ReturnArray, 2)) = std_Chemistry_GetPSE(ElementIndex, ReturnWhat)
                            ReturnArray(1, UBound(ReturnArray, 2)) = 1
                            ReDim Preserve ReturnArray(1, UBound(ReturnArray, 2) + 1)
                        End If
                        Success = True
                    End If
            End Select
            If Success = True Then
                StartPos = StartPos + EndPos
                EndPos = Len(Text) - StartPos + 1
                Success = False
                GoTo Skip
            Else
                EndPos = EndPos -1
                If EndPos < 0 Then Exit Function
            End If
        Loop
        Skip:
    Next i
    
    For i = 0 To MaxDepth
        For j = 0 To Ubound(Temp, 3) - 1
            Found = False
            For k = 0 To UBound(ReturnArray, 2) - 1
                If Temp(0, i, j) = ReturnArray(0, k) Then
                    Found = True
                    ReturnArray(1, k) = ReturnArray(1, k) + Temp(1, i, j)
                End If
            Next k
            If Found = False Then
                ReDim Preserve ReturnArray(1, UBound(ReturnArray, 2) + 1)
                ReturnArray(0, k) = Temp(0, i, j)
                ReturnArray(1, k) = Temp(1, i, j)
            End If
        Next j
    Next i
    std_Chemistry_Intepret = ReturnArray

End Function

Public Function std_Chemistry_SearchPSE(Value As Variant, SearchWhat As std_Chemistry_Element_Property) As Long
    Dim i As Long
    Call Initialize()
    For i = 0 To 118
        Select Case SearchWhat
            Case std_Chemistry_Element_Property.AtomicNumber : If CLng(Value) = PSE(i).AtomicNumber Then std_Chemistry_SearchPSE = i: Exit Function
            Case std_Chemistry_Element_Property.Name         : If CStr(Value) = PSE(i).Name         Then std_Chemistry_SearchPSE = i: Exit Function
            Case std_Chemistry_Element_Property.Short        : If CStr(Value) = PSE(i).Short        Then std_Chemistry_SearchPSE = i: Exit Function
            Case std_Chemistry_Element_Property.MolarMass    : If CDbl(Value) = PSE(i).MolarMass    Then std_Chemistry_SearchPSE = i: Exit Function
        End Select
    Next i
End Function

Public Function std_Chemistry_GetPSE(AtomicNumber As Long, SearchWhat As std_Chemistry_Element_Property) As Variant
    Call Initialize()
    Select Case SearchWhat
        Case std_Chemistry_Element_Property.AtomicNumber : std_Chemistry_GetPSE = PSE(AtomicNumber).AtomicNumber
        Case std_Chemistry_Element_Property.Name         : std_Chemistry_GetPSE = PSE(AtomicNumber).Name
        Case std_Chemistry_Element_Property.Short        : std_Chemistry_GetPSE = PSE(AtomicNumber).Short
        Case std_Chemistry_Element_Property.MolarMass    : std_Chemistry_GetPSE = PSE(AtomicNumber).MolarMass
    End Select
End Function

Public Function std_Chemistry_GetMolarMass(Arr() As Variant) As Double
    Dim i As Long
    Call Initialize()
    For i = 0 To Ubound(Arr, 2)
        std_Chemistry_GetMolarMass = std_Chemistry_GetMolarMass + (Arr(0, i) * Arr(1, i))
    Next i
End Function

Private Sub Initialize()
    Static Initialized As Boolean
    If Initialized Then Exit Sub
    PSE(001).AtomicNumber = 001: PSE(001).Name = "Hydrogen"      : PSE(001).Short = "H"  : PSE(001).MolarMass = 01.007825
    PSE(002).AtomicNumber = 002: PSE(002).Name = "Helium"        : PSE(002).Short = "He" : PSE(002).MolarMass = 04.002603
    PSE(003).AtomicNumber = 003: PSE(003).Name = "Lithium"       : PSE(003).Short = "Li" : PSE(003).MolarMass = 06.94
    PSE(004).AtomicNumber = 004: PSE(004).Name = "Beryllium"     : PSE(004).Short = "Be" : PSE(004).MolarMass = 09.012
    PSE(005).AtomicNumber = 005: PSE(005).Name = "Boron"         : PSE(005).Short = "B"  : PSE(005).MolarMass = 10.81
    PSE(006).AtomicNumber = 006: PSE(006).Name = "Carbon"        : PSE(006).Short = "C"  : PSE(006).MolarMass = 12.011
    PSE(007).AtomicNumber = 007: PSE(007).Name = "Nitrogen"      : PSE(007).Short = "N"  : PSE(007).MolarMass = 14.007
    PSE(008).AtomicNumber = 008: PSE(008).Name = "Oxygen"        : PSE(008).Short = "O"  : PSE(008).MolarMass = 15.999
    PSE(009).AtomicNumber = 009: PSE(009).Name = "Fluorine"      : PSE(009).Short = "F"  : PSE(009).MolarMass = 18.998
    PSE(010).AtomicNumber = 010: PSE(010).Name = "Neon"          : PSE(010).Short = "Ne" : PSE(010).MolarMass = 20.1797
    PSE(011).AtomicNumber = 011: PSE(011).Name = "Sodium"        : PSE(011).Short = "Na" : PSE(011).MolarMass = 22.989
    PSE(012).AtomicNumber = 012: PSE(012).Name = "Magnesium"     : PSE(012).Short = "Mg" : PSE(012).MolarMass = 24.305
    PSE(013).AtomicNumber = 013: PSE(013).Name = "Aluminium"     : PSE(013).Short = "Al" : PSE(013).MolarMass = 26.981
    PSE(014).AtomicNumber = 014: PSE(014).Name = "Silicon"       : PSE(014).Short = "Si" : PSE(014).MolarMass = 28.085
    PSE(015).AtomicNumber = 015: PSE(015).Name = "Phosphorus"    : PSE(015).Short = "P"  : PSE(015).MolarMass = 30.973
    PSE(016).AtomicNumber = 016: PSE(016).Name = "Sulfur"        : PSE(016).Short = "S"  : PSE(016).MolarMass = 32.06
    PSE(017).AtomicNumber = 017: PSE(017).Name = "Chlorine"      : PSE(017).Short = "Cl" : PSE(017).MolarMass = 35.45
    PSE(018).AtomicNumber = 018: PSE(018).Name = "Argon"         : PSE(018).Short = "Ar" : PSE(018).MolarMass = 39.95
    PSE(019).AtomicNumber = 019: PSE(019).Name = "Potassium"     : PSE(019).Short = "K"  : PSE(019).MolarMass = 39.0983
    PSE(020).AtomicNumber = 020: PSE(020).Name = "Calcium"       : PSE(020).Short = "Ca" : PSE(020).MolarMass = 40.078
    PSE(021).AtomicNumber = 021: PSE(021).Name = "Scandium"      : PSE(021).Short = "Sc" : PSE(021).MolarMass = 44.955
    PSE(022).AtomicNumber = 022: PSE(022).Name = "Titanium"      : PSE(022).Short = "Ti" : PSE(022).MolarMass = 47.867
    PSE(023).AtomicNumber = 023: PSE(023).Name = "Vanadium"      : PSE(023).Short = "V"  : PSE(023).MolarMass = 50.9415
    PSE(024).AtomicNumber = 024: PSE(024).Name = "Chromium"      : PSE(024).Short = "Cr" : PSE(024).MolarMass = 51.9961
    PSE(025).AtomicNumber = 025: PSE(025).Name = "Manganese"     : PSE(025).Short = "Mn" : PSE(025).MolarMass = 54.938
    PSE(026).AtomicNumber = 026: PSE(026).Name = "Iron"          : PSE(026).Short = "Fe" : PSE(026).MolarMass = 55.845
    PSE(027).AtomicNumber = 027: PSE(027).Name = "Cobalt"        : PSE(027).Short = "Co" : PSE(027).MolarMass = 58.933
    PSE(028).AtomicNumber = 028: PSE(028).Name = "Nickel"        : PSE(028).Short = "Ni" : PSE(028).MolarMass = 58.6934
    PSE(029).AtomicNumber = 029: PSE(029).Name = "Copper"        : PSE(029).Short = "Cu" : PSE(029).MolarMass = 63.546
    PSE(030).AtomicNumber = 030: PSE(030).Name = "Zinc"          : PSE(030).Short = "Zn" : PSE(030).MolarMass = 65.38
    PSE(031).AtomicNumber = 031: PSE(031).Name = "Gallium"       : PSE(031).Short = "Ga" : PSE(031).MolarMass = 69.723
    PSE(032).AtomicNumber = 032: PSE(032).Name = "Germanium"     : PSE(032).Short = "Ge" : PSE(032).MolarMass = 72.63
    PSE(033).AtomicNumber = 033: PSE(033).Name = "Arsenic"       : PSE(033).Short = "As" : PSE(033).MolarMass = 74.921
    PSE(034).AtomicNumber = 034: PSE(034).Name = "Selenium"      : PSE(034).Short = "Se" : PSE(034).MolarMass = 78.971
    PSE(035).AtomicNumber = 035: PSE(035).Name = "Bromine"       : PSE(035).Short = "Br" : PSE(035).MolarMass = 79.904
    PSE(036).AtomicNumber = 036: PSE(036).Name = "Krypton"       : PSE(036).Short = "Kr" : PSE(036).MolarMass = 83.798
    PSE(037).AtomicNumber = 037: PSE(037).Name = "Rubidium"      : PSE(037).Short = "Rb" : PSE(037).MolarMass = 85.4678
    PSE(038).AtomicNumber = 038: PSE(038).Name = "Strontium"     : PSE(038).Short = "Sr" : PSE(038).MolarMass = 87.62
    PSE(039).AtomicNumber = 039: PSE(039).Name = "Yttrium"       : PSE(039).Short = "Y"  : PSE(039).MolarMass = 88.905
    PSE(040).AtomicNumber = 040: PSE(040).Name = "Zirconium"     : PSE(040).Short = "Zr" : PSE(040).MolarMass = 91.224
    PSE(041).AtomicNumber = 041: PSE(041).Name = "Niobium"       : PSE(041).Short = "Nb" : PSE(041).MolarMass = 92.906
    PSE(042).AtomicNumber = 042: PSE(042).Name = "Molybdenum"    : PSE(042).Short = "Mo" : PSE(042).MolarMass = 95.95
    PSE(043).AtomicNumber = 043: PSE(043).Name = "Technetium"    : PSE(043).Short = "Tc" : PSE(043).MolarMass = 97
    PSE(044).AtomicNumber = 044: PSE(044).Name = "Ruthenium"     : PSE(044).Short = "Ru" : PSE(044).MolarMass = 101.07
    PSE(045).AtomicNumber = 045: PSE(045).Name = "Rhodium"       : PSE(045).Short = "Rh" : PSE(045).MolarMass = 102.905
    PSE(046).AtomicNumber = 046: PSE(046).Name = "Palladium"     : PSE(046).Short = "Pd" : PSE(046).MolarMass = 106.42
    PSE(047).AtomicNumber = 047: PSE(047).Name = "Silver"        : PSE(047).Short = "Ag" : PSE(047).MolarMass = 107.8682
    PSE(048).AtomicNumber = 048: PSE(048).Name = "Cadmium"       : PSE(048).Short = "Cd" : PSE(048).MolarMass = 112.414
    PSE(049).AtomicNumber = 049: PSE(049).Name = "Indium"        : PSE(049).Short = "In" : PSE(049).MolarMass = 114.818
    PSE(050).AtomicNumber = 050: PSE(050).Name = "Tin"           : PSE(050).Short = "Sn" : PSE(050).MolarMass = 118.71
    PSE(051).AtomicNumber = 051: PSE(051).Name = "Antimony"      : PSE(051).Short = "Sb" : PSE(051).MolarMass = 121.76
    PSE(052).AtomicNumber = 052: PSE(052).Name = "Tellerium"     : PSE(052).Short = "Te" : PSE(052).MolarMass = 127.6
    PSE(053).AtomicNumber = 053: PSE(053).Name = "Iodine"        : PSE(053).Short = "I"  : PSE(053).MolarMass = 126.904
    PSE(054).AtomicNumber = 054: PSE(054).Name = "Xenon"         : PSE(054).Short = "Xe" : PSE(054).MolarMass = 131.293
    PSE(055).AtomicNumber = 055: PSE(055).Name = "Cesium"        : PSE(055).Short = "Cs" : PSE(055).MolarMass = 132.905
    PSE(056).AtomicNumber = 056: PSE(056).Name = "Barium"        : PSE(056).Short = "Ba" : PSE(056).MolarMass = 137.327
    PSE(057).AtomicNumber = 057: PSE(057).Name = "Lanthanium"    : PSE(057).Short = "La" : PSE(057).MolarMass = 138.905
    PSE(058).AtomicNumber = 058: PSE(058).Name = "Cerium"        : PSE(058).Short = "Ce" : PSE(058).MolarMass = 140.116
    PSE(059).AtomicNumber = 059: PSE(059).Name = "Praseodymium"  : PSE(059).Short = "Pr" : PSE(059).MolarMass = 140.907
    PSE(060).AtomicNumber = 060: PSE(060).Name = "Neodymium"     : PSE(060).Short = "Nd" : PSE(060).MolarMass = 144.242
    PSE(061).AtomicNumber = 061: PSE(061).Name = "Promethium"    : PSE(061).Short = "Pm" : PSE(061).MolarMass = 145
    PSE(062).AtomicNumber = 062: PSE(062).Name = "Samarium"      : PSE(062).Short = "Sm" : PSE(062).MolarMass = 150.36
    PSE(063).AtomicNumber = 063: PSE(063).Name = "Europium"      : PSE(063).Short = "Eu" : PSE(063).MolarMass = 151.964
    PSE(064).AtomicNumber = 064: PSE(064).Name = "Gadolinium"    : PSE(064).Short = "Gd" : PSE(064).MolarMass = 157.25
    PSE(065).AtomicNumber = 065: PSE(065).Name = "Terbium"       : PSE(065).Short = "Tb" : PSE(065).MolarMass = 158.925
    PSE(066).AtomicNumber = 066: PSE(066).Name = "Dysprosium"    : PSE(066).Short = "Dy" : PSE(066).MolarMass = 162.5
    PSE(067).AtomicNumber = 067: PSE(067).Name = "Holmium"       : PSE(067).Short = "Ho" : PSE(067).MolarMass = 164.93
    PSE(068).AtomicNumber = 068: PSE(068).Name = "Ervuzn"        : PSE(068).Short = "Er" : PSE(068).MolarMass = 167.259
    PSE(069).AtomicNumber = 069: PSE(069).Name = "Thulium"       : PSE(069).Short = "Tm" : PSE(069).MolarMass = 168.934
    PSE(070).AtomicNumber = 070: PSE(070).Name = "Ytterbium"     : PSE(070).Short = "Yb" : PSE(070).MolarMass = 173.045
    PSE(071).AtomicNumber = 071: PSE(071).Name = "Lutetium"      : PSE(071).Short = "Lu" : PSE(071).MolarMass = 174.9668
    PSE(072).AtomicNumber = 072: PSE(072).Name = "Hafnium"       : PSE(072).Short = "Hf" : PSE(072).MolarMass = 178.486
    PSE(073).AtomicNumber = 073: PSE(073).Name = "Tantalum"      : PSE(073).Short = "Ta" : PSE(073).MolarMass = 180.947
    PSE(074).AtomicNumber = 074: PSE(074).Name = "Tungsten"      : PSE(074).Short = "W"  : PSE(074).MolarMass = 183.84
    PSE(075).AtomicNumber = 075: PSE(075).Name = "Rhenium"       : PSE(075).Short = "Re" : PSE(075).MolarMass = 186.207
    PSE(076).AtomicNumber = 076: PSE(076).Name = "Osmium"        : PSE(076).Short = "Os" : PSE(076).MolarMass = 190.23
    PSE(077).AtomicNumber = 077: PSE(077).Name = "Iridium"       : PSE(077).Short = "Ir" : PSE(077).MolarMass = 192.217
    PSE(078).AtomicNumber = 078: PSE(078).Name = "Platinum"      : PSE(078).Short = "Pt" : PSE(078).MolarMass = 195.084
    PSE(079).AtomicNumber = 079: PSE(079).Name = "Gold"          : PSE(079).Short = "Au" : PSE(079).MolarMass = 196.966
    PSE(080).AtomicNumber = 080: PSE(080).Name = "Mercury"       : PSE(080).Short = "Hg" : PSE(080).MolarMass = 200.592
    PSE(081).AtomicNumber = 081: PSE(081).Name = "Thallium"      : PSE(081).Short = "Ti" : PSE(081).MolarMass = 204.38
    PSE(082).AtomicNumber = 082: PSE(082).Name = "Lead"          : PSE(082).Short = "Pb" : PSE(082).MolarMass = 207.2
    PSE(083).AtomicNumber = 083: PSE(083).Name = "Bismuth"       : PSE(083).Short = "Bi" : PSE(083).MolarMass = 208.98
    PSE(084).AtomicNumber = 084: PSE(084).Name = "Polonium"      : PSE(084).Short = "Po" : PSE(084).MolarMass = 209
    PSE(085).AtomicNumber = 085: PSE(085).Name = "Astatine"      : PSE(085).Short = "At" : PSE(085).MolarMass = 210
    PSE(086).AtomicNumber = 086: PSE(086).Name = "Radon"         : PSE(086).Short = "Rn" : PSE(086).MolarMass = 222
    PSE(087).AtomicNumber = 087: PSE(087).Name = "Francium"      : PSE(087).Short = "Fr" : PSE(087).MolarMass = 223
    PSE(088).AtomicNumber = 088: PSE(088).Name = "Radium"        : PSE(088).Short = "Ra" : PSE(088).MolarMass = 226
    PSE(089).AtomicNumber = 089: PSE(089).Name = "Actinium"      : PSE(089).Short = "Ac" : PSE(089).MolarMass = 227
    PSE(090).AtomicNumber = 090: PSE(090).Name = "Thorium"       : PSE(090).Short = "Th" : PSE(090).MolarMass = 232.0377
    PSE(091).AtomicNumber = 091: PSE(091).Name = "Protactinium"  : PSE(091).Short = "Pa" : PSE(091).MolarMass = 231.035
    PSE(092).AtomicNumber = 092: PSE(092).Name = "Uranium"       : PSE(092).Short = "U"  : PSE(092).MolarMass = 238.028
    PSE(093).AtomicNumber = 093: PSE(093).Name = "Neptunium"     : PSE(093).Short = "Np" : PSE(093).MolarMass = 237
    PSE(094).AtomicNumber = 094: PSE(094).Name = "Plutonium"     : PSE(094).Short = "Pu" : PSE(094).MolarMass = 244
    PSE(095).AtomicNumber = 095: PSE(095).Name = "Americium"     : PSE(095).Short = "Am" : PSE(095).MolarMass = 243
    PSE(096).AtomicNumber = 096: PSE(096).Name = "Curium"        : PSE(096).Short = "Cm" : PSE(096).MolarMass = 247
    PSE(097).AtomicNumber = 097: PSE(097).Name = "Berkelium"     : PSE(097).Short = "Bk" : PSE(097).MolarMass = 247
    PSE(098).AtomicNumber = 098: PSE(098).Name = "Californium"   : PSE(098).Short = "Cf" : PSE(098).MolarMass = 251
    PSE(099).AtomicNumber = 099: PSE(099).Name = "Einsteinium"   : PSE(099).Short = "Es" : PSE(099).MolarMass = 252
    PSE(100).AtomicNumber = 100: PSE(100).Name = "Fermium"       : PSE(100).Short = "Fm" : PSE(100).MolarMass = 257
    PSE(101).AtomicNumber = 101: PSE(101).Name = "Mendelevium"   : PSE(101).Short = "Md" : PSE(101).MolarMass = 258
    PSE(102).AtomicNumber = 102: PSE(102).Name = "Nobelium"      : PSE(102).Short = "No" : PSE(102).MolarMass = 259
    PSE(103).AtomicNumber = 103: PSE(103).Name = "Lawrencium"    : PSE(103).Short = "Lr" : PSE(103).MolarMass = 262
    PSE(104).AtomicNumber = 104: PSE(104).Name = "Rutherfordium" : PSE(104).Short = "Rf" : PSE(104).MolarMass = 267
    PSE(105).AtomicNumber = 105: PSE(105).Name = "Dubnium"       : PSE(105).Short = "Db" : PSE(105).MolarMass = 270
    PSE(106).AtomicNumber = 106: PSE(106).Name = "Seaborgium"    : PSE(106).Short = "Sg" : PSE(106).MolarMass = 269
    PSE(107).AtomicNumber = 107: PSE(107).Name = "Bohrium"       : PSE(107).Short = "Bh" : PSE(107).MolarMass = 270
    PSE(108).AtomicNumber = 108: PSE(108).Name = "Hassium"       : PSE(108).Short = "Hs" : PSE(108).MolarMass = 270
    PSE(109).AtomicNumber = 109: PSE(109).Name = "Meitnerium"    : PSE(109).Short = "Mt" : PSE(109).MolarMass = 278
    PSE(110).AtomicNumber = 110: PSE(110).Name = "Darmstadtium"  : PSE(110).Short = "Ds" : PSE(110).MolarMass = 281
    PSE(111).AtomicNumber = 111: PSE(111).Name = "Roentgenium"   : PSE(111).Short = "Rg" : PSE(111).MolarMass = 281
    PSE(112).AtomicNumber = 112: PSE(112).Name = "Copernicium"   : PSE(112).Short = "Cn" : PSE(112).MolarMass = 285
    PSE(113).AtomicNumber = 113: PSE(113).Name = "Nihonium"      : PSE(113).Short = "Nh" : PSE(113).MolarMass = 286
    PSE(114).AtomicNumber = 114: PSE(114).Name = "Flerovium"     : PSE(114).Short = "Fl" : PSE(114).MolarMass = 289
    PSE(115).AtomicNumber = 115: PSE(115).Name = "Moscovium"     : PSE(115).Short = "Mc" : PSE(115).MolarMass = 289
    PSE(116).AtomicNumber = 116: PSE(116).Name = "Livermorium"   : PSE(116).Short = "Lv" : PSE(116).MolarMass = 293
    PSE(117).AtomicNumber = 117: PSE(117).Name = "Tenness"       : PSE(117).Short = "Ts" : PSE(117).MolarMass = 293
    PSE(118).AtomicNumber = 118: PSE(118).Name = "Oganesson"     : PSE(118).Short = "Og" : PSE(118).MolarMass = 294
    Initialized = True
End Sub