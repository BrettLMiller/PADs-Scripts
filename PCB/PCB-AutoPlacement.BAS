' PCB-AutoPlacement.bas   used with PCB only

' Author: BL Miller
' This work is licensed under the Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License.
' To view a copy of this license, visit http://creativecommons.org/licenses/by-nc-sa/4.0/ or send a letter to Creative Commons, PO Box 1866, Mountain View, CA 94042, USA.

' Usage: (PCB)
' Requires part attributes block & cell & master cell info (block,cell,master_des) to position components.
' uses XY & rotation from master cell
' Uses Board attribute "Placement.TestJig" to get geometry
' Any "cell" can be master
' Origin does not matter.

' 20141125  BLM  Added Component Ref.Des. label position placement (matches master)
' 20150219  BLM  Test X&Y postion etc before moving (speed up).
'                Added label text height & width, first testing if changes needed.
'                Set dimension units const ppcbUnits to match the Grid64Data array. PCB job can be (left) in any "units".
' 20150226  BLM  block repeat X axis error was 2500 should be 2300
' 20150514  BLM  Added transform & offset array to support mercury (no mirror symmetry)
'                               Changed attributes used for placement info
' 20150520  BLM  Cell placement geometry type text required to be loaded in PCB doc Attribute "Placement.TestJig"
' 20150528  BLM  Hexadecimal attribute values for "CELL NUM" need prefix of "&H" .. just use simple decimal numbers.
' 201511--  BLM  added OCXO Mercury Aging geometry
' 20161004  BLM  added TCOXCO geometry
' 20161006  BLM  warn skip glued components.

Attribute VB_Name = "Module1"
Option Explicit

Const envTemp As String =   "TEMP"
Const temppath As String =  "c:/temp/"
Const PLogicTxt As String = "PADS Logic"
Const PPCBTxt As String =   "PowerPCB"

Const MasterBlock As String = "A"                        ' set the master cell "block" value
Const MasterCell As String =  "0"                        ' set the master cell "cell" value  use "&HA" for 10 in hexadecimal

Const BO_Block As String =      "CELL BLOCK"             ' attribute name for block value
Const BO_Cell As String =       "CELL NUM"
Const BO_MasterDesg As String = "CELL MDES"

Const PCBPlacementAttr As String = "Placement.TestJig"   ' PCB attribute used to set the placement format

Dim MasterParts As objects
Dim parts As objects
Dim AttrList As Attributes
Dim MasAttrList As Attributes
Dim messageTxt As String
Dim MsgAnswer As Long
Dim offsetX As Double, offsetY As Double
Dim starttime As String
Dim blnQuestionableData As Boolean

'  position = GridData * {transform} + {offset}
Dim GridData As Variant
Dim GridOffset As Variant               ' block offset matrix { A B C D}
Dim GridSymmetry As Variant             ' transform matrix { A B C D}
' Dim Block_offsetX As Double 
Dim ppcbUnits As Long                   ' ppcbUnitCurrent, ppcbUnitDatabase, ppcbUnitInch, ppcbUnitMetric
Dim TestCardType As Long

' std LFM pallet/testcard 4 blocks by 16 DUTs/block
' matrix maths: B = A * {1, -1} + {0, 0} ; C = A + {2300, 0} ; D = B + {2300, 0}
' cell 0 to F(15) as (X,Y) offsets from origin in centre (mils)
' Const Grid64Block_offsetX As Double = 2300  old code
' GridnnData()   - all relative positions.
' does not really matter where the origin is just make it easy to work out the numbers
' below GridnnData() originated from drawing with origin in the middle

Const Grid64Units As Long = ppcbUnitMils
'---------------------------------------------------------------------------------------------------------------------
' std LFM testcard 4 block of 16
Const Grid64Symmetry = Array(Array(1, 1), Array(1, -1), Array(1, 1), Array(1, -1) )
Const Grid64Offset = Array(Array(0, 0),Array(0, 0),Array(2300, 0),Array(2300, 0) )
Const Grid64Data = Array(Array(-1900,-400), Array(-1900,-750), Array(-1900,-1100), Array(-1900,-1450), _
                         Array(-1400,-400), Array(-1400,-750), Array(-1400,-1100), Array(-1400,-1450), _
                         Array(-900,-400),  Array(-900,-750),  Array(-900,-1100),  Array(-900,-1450), _
                         Array(-400,-400),  Array(-400,-750),  Array(-400,-1100),  Array(-400,-1450), _
                        )
Const LFM16x4 As Long = 1640
'---------------------------------------------------------------------------------------------------------------------
' std LFM Mercury testcard  4 blocks 10DUTs/block; no mirror x axis symmetry; all direct offsets & metric
' B = {0, 47} + {1, 1} * A ; C = {60.5, 0} + {1, 1} * A ; D=  {60.5, 47} + {1, 1} * A
' cell 0 to 9 as (x,y) offsets from origin as centre PCB.
Const MercuryGrid40Units As Long =  ppcbUnitMetric
Const MercuryGrid40Symmetry = Array(Array(1, 1), Array(1, 1), Array(1, 1), Array(1, 1) )
Const MercuryGrid40Offset = Array(Array(0, 0), Array(0, 47), Array(60.5, 0), Array(60.5, 47) )
Const MercuryGrid40Data = Array(Array(-53.25, -13.5), Array(-53.25, -33.5), Array(-41.75, -13.5), Array(-41.75, -33.5), _
                                Array(-30.25, -13.5), Array(-30.25, -33.5), Array(-18.75, -13.5), Array(-18.75, -33.5), _
                                Array( -7.25, -13.5), Array( -7.25, -33.5)  )     'Array(0,0),  Array(0,0), _
                                'Array(0,0),  Array(0,0),  Array(0,0),  Array(0,0),  )
Const Mercury10x4 As Long = 1040
Const txtMercury10x4 As String = "Mercury40"
'---------------------------------------------------------------------------------------------------------------------
' OCXO Mercury Ageing System card 4 wide by 15 long; cells 30mm x 50mm
Const OCXOMercuryGrid60Units As Long =  ppcbUnitMetric
Const OCXOMercuryGrid60Symmetry = Array(Array(1, 1), Array(1, 1), Array(1, 1), Array(1, 1) )
Const OCXOMercuryGrid60Offset = Array(Array(0, 0), Array(0, 50), Array(0, -50), Array(0, 100) )
Const OCXOMercuryGrid60Data = Array(Array(0, 0), Array(30, 0), Array(60, 0), Array(90, 0), _
                                Array(120, 0), Array(150, 0), Array(180, 0), Array(210, 0), _
                                Array( 240, 0), Array( 270, 0), Array(300,0),  Array(330,0), _
                                Array(360,0),  Array(390,0),  Array(420,0), Array(450,0) )   ',  Array(0,0),  )
Const OCXOMercury60 As Long = 1540
Const txtOCXOMercury60 As String = "OCXO-M60"
'---------------------------------------------------------------------------------------------------------------------
' OCXO Mercury Ageing System card 4 wide by 4 by 4 long (4 blocks of 16); cells 30mm x 50mm
Const OCXOMercuryGrid64Units As Long =  ppcbUnitMetric
Const OCXOMercuryGrid64Symmetry = Array(Array(1, 1), Array(1, 1), Array(1, 1), Array(1, 1) )
Const OCXOMercuryGrid64Offset = Array(Array(0, 0), Array(120, 0), Array(240, 0), Array(360,  0) )
Const OCXOMercuryGrid64Data = Array(Array( 0, 150), Array( 0, 100), Array( 0, 50), Array( 0, 0), _
                                    Array(30, 150), Array(30, 100), Array(30, 50), Array(30, 0), _
                                    Array(60, 150), Array(60, 100), Array(60, 50), Array(60, 0), _
                                    Array(90, 150), Array(90, 100), Array(90, 50), Array(90, 0) )   ',  Array(0,0),  )
Const OCXOMercury64 As Long = 1644
Const txtOCXOMercury64 As String = "OCXO-M64"
'---------------------------------------------------------------------------------------------------------------------
' TCOCXO System card 4 blocks wide by 6 long (4 blocks of 6); cells 50mm x 60mm
Const TCOCXOGrid24Units As Long =  ppcbUnitMetric
Const TCOCXOGrid24Symmetry = Array(Array(1, 1), Array(1, 1), Array(1, 1), Array(1, 1) )
Const TCOCXOGrid24Offset  = Array(Array(0, 0), Array(0, 60), Array(0, 120), Array(0,  180) )
Const TCOCXOGrid24Data = Array(Array(   0, 0), Array( 50, 0), Array(100, 0), _
                               Array( 150, 0), Array(200, 0), Array(250, 0) )    ', Array(30, 50), Array(30, 0), _
                                   ' Array(60, 150), Array(60, 100), Array(60, 50), Array(60, 0), _
                                   ' Array(90, 150), Array(90, 100), Array(90, 50), Array(90, 0) )   ',  Array(0,0),  )
Const TCOCXO24 As Long = 6424
Const txtTCOCXO24 As String = "TCOCXO_24"
'---------------------------------------------------------------------------------------------------------------------

Const txtVCXO_8x4 As String = "VCXO_7050_32"
Const VCXO8x4 As Long = 705032
Const txtVCXO_Panel As String = "VCXO_1510_40"
Const VCXO8x5 As Long = 151040

Sub main
    Dim HostAppName As String
    Dim wintemppath As String
    Dim selectcomp As Long
    Dim part As Component
    Dim mpart As Component
    Dim partlabels As objects, plabel As label
    Dim mpartlabels As objects, mplabel As label
    Dim found As Boolean, lfound As Boolean, skip As Boolean
    Dim MP_offsetX As Double, MP_offsetY As Double, MP_Rot As Double, MP_layer As Long
    Dim tempX As Double, tempY As Double
    Dim strPlacementTypeTxt As String

    blnQuestionableData = False
    strPlacementTypeTxt = "default LFM 16 x 4"
    TestCardType = LFM16x4                          ' default cell placement geometry
    If Not Document.Attributes(PCBPlacementAttr) Is Nothing Then
        strPlacementTypeTxt = Document.Attributes(PCBPlacementAttr)
        Debug.Print strPlacementTypeTxt

        Select Case strPlacementTypeTxt
        Case txtOCXOMercury60
            TestcardType = OCXOMercury60
        Case txtOCXOMercury64
            TestcardType = OCXOMercury64
        Case txtMercury10x4
            TestcardType = Mercury10x4
        Case txtVCXO_Panel
            TestcardType = VCXO8x5
        Case txtVCXO_8x4
            TestcardType = VCXO8x4
        case txtTCOCXO24
            TestcardType = TCOCXO24
        End Select
    End If
     
    ppcbUnits = ppcbUnitMils
    Select Case TestcardType
    Case Mercury10x4
        GridData =     MercuryGrid40Data
        GridOffset =   MercuryGrid40Offset
        GridSymmetry = MercuryGrid40Symmetry
        ppcbUnits =    MercuryGrid40Units
    Case OCXOMercury60
        GridData =     OCXOMercuryGrid60Data
        GridOffset =   OCXOMercuryGrid60Offset
        GridSymmetry = OCXOMercuryGrid60Symmetry
        ppcbUnits =    OCXOMercuryGrid60Units
    Case OCXOMercury64
        GridData =     OCXOMercuryGrid64Data
        GridOffset =   OCXOMercuryGrid64Offset
        GridSymmetry = OCXOMercuryGrid64Symmetry
        ppcbUnits =    OCXOMercuryGrid64Units
    case TCOCXO24
        GridData =     TCOCXOGrid24Data
        GridOffset =   TCOCXOGrid24Offset  
        GridSymmetry = TCOCXOGrid24Symmetry
        ppcbUnits =    TCOCXOGrid24Units
    Case Else
        GridData =     Grid64Data
        GridOffset =   Grid64Offset
        GridSymmetry = Grid64Symmetry
        ppcbUnits =    Grid64Units
    End Select
    
    HostAppName = Application.Name
    wintemppath = Environ(envTemp)
    If wintemppath = "" Then wintemppath = temppath
'    If HostAppName = "PLogicTxt" Then
'       Dim ppcbLibraryItemTypePartType As Long
'    End If 
'    If HostAppName = "PPCBTxt" Then
'       Dim plogObjectTypeComponent As Long
'    End If 
      If HostAppName = PLogicTxt Then
         selectcomp = 1&     'plogObjectTypeComponent
      Else
         selectcomp = 2&     'ppcbLibraryItemTypePartType
      End If
      
      Debug.Print Now
      
' find all the master block/cell parts to get XYRS & desg. 
    ActiveDocument.SelectObjects(,,False)

    Set parts = Document.Components
	For Each part In parts
        Set AttrList = part.Attributes
	    If Not AttrList(BO_Block) Is Nothing Then
	        If Not AttrList(BO_Cell) Is Nothing Then
	            If Not AttrList(BO_MasterDesg) Is Nothing Then
	        
	                If AttrList(BO_Block) = MasterBlock And AttrList(BO_Cell) = MasterCell And AttrList(BO_MasterDesg) <> "" Then
	                    part.selected = True
	                    'If MasterParts Is Nothing Then
	                    '    Set MasterParts = part
	                    'Else
	                    '    MasterParts.Add(part)
	                    'End If
                    End If
                End If
            End If
	    End If
	Next part
	
	Set MasterParts = ActiveDocument.GetObjects(ppcbObjectTypeComponent, "*", True) 
	Debug.Print MasterParts.Count
	messageTxt = "Master Part Count = " & Str(MasterParts.Count)
	messageTxt = messageTxt & Chr$(13) & "PCB Placement Type " & strPlacementTypeTxt
		
	MsgAnswer = MsgBox(messageTxt, vbYesNoCancel)
	
	' search all parts skipping masters
    If MsgAnswer = vbYes And MasterParts.Count > 0 Then
	    For Each part In parts
	        Set AttrList = part.Attributes
            found = False
            ' check part is not in master parts collection
            If MasterParts(part.Name) Is Nothing Then
                If Not AttrList(BO_Block) Is Nothing And Not AttrList(BO_Cell) Is Nothing _
	               And Not AttrList(BO_MasterDesg) Is Nothing Then
                    skip = True         ' do no parts

                    ' option hack: place only designators starting with--
                    'If Left(part.Name,3) = "C17" Then skip = False
                    If Left(part.Name,2) = "TJ" Then skip = False
                    If Left(part.Name,1) = "Q" Then skip = False
                    
                    ' option hack: place parts with match to master designator
                    'If AttrList(BO_MasterDesg) = "C1" Then skip = False
                    'If AttrList(BO_MasterDesg) = "U2" Then skip = False
                    'If AttrList(BO_MasterDesg) = "C4" Then skip = False
                    'If AttrList(BO_MasterDesg) = "D4" Then skip = False
                    'If AttrList(BO_MasterDesg) = "C17" Then skip = False
                    'If Left(part.Name,1) = "ASS_SPL" Then skip = False
                    
                    ' option hack: place only block x
                    'If AttrList(BO_Block) = "A" Then  
                        If  AttrList(BO_Cell) = "0" Then
                            skip = False
                        End If
                    'End If

                    ' option do all parts
                     skip = False

                    ' option hack leave "placed" parts alone
                    'If part.Placed Then skip = True   ' TBD

                    If Len(AttrList(BO_Block)) < 1 Then skip = True
                    If InStr(1, "ABCD", AttrList(BO_Block)) < 1 Then skip = True
                    If AttrList(BO_Cell) = "" Then skip = True
                    If AttrList(BO_Cell) = " " Then skip = True
                    If Not skip Then
	                    ' find match designator in master parts
                        For Each mpart In MasterParts
                            Set MasAttrList = mpart.Attributes

                            If AttrList(BO_MasterDesg) = MasAttrList(BO_MasterDesg) Then
	                            Debug.Print part.Name;" : ";mpart.Name;
	                            found = True
                                'calculate master part cell offsets X & Y from master cell offset info
	                            Call CalcOffsets(MasAttrList(BO_Block), MasAttrList(BO_Cell))
	                            MP_offsetX = mpart.PositionX(ppcbUnits) - offsetX :     MP_offsetY = mpart.PositionY(ppcbUnits) - offsetY   ' centerX centerY  ??? MoveCenter() ???
	                            MP_Rot = mpart.Orientation : MP_Layer = mpart.layer

                                'calculate part cell offsets X & Y from its own block cell BO info
	                            Call CalcOffsets(AttrList(BO_Block), AttrList(BO_Cell))

                                offsetX = offsetX + MP_offsetX: offsetY = offsetY + MP_offsetY
	                        
                                Debug.Print offsetX; " ";offsetY; " O"; MP_Rot; " L"; MP_Layer
                                If part.Glued Then
                                   Debug.Print part.Name;" : glued";
                                Else   
                                    If part.PositionX(ppcbUnits) <> offsetX Or part.PositionY(ppcbUnits) <> offsetY Then
                                        part.Move(offsetX, offsetY,(ppcbUnits))
                                    End If
                                    If part.Orientation <> MP_Rot Then part.Orientation = MP_Rot
                                    If part.layer <> MP_Layer Then part.layer = MP_Layer
                                End If
                                ' Ref.Des. labels
                                lfound = False
                                Set mpartlabels = mpart.Labels
                                For Each mplabel In mpartlabels
                                    If mplabel.LabelType = ppcbLabelTypeRefDesignator Then
                                        lfound = True
                                        Exit For
                                    End If
                                Next mplabel
                                
                                ' repos the part label
                                If lfound Then
                                    Set partlabels = part.Labels
                                    For Each plabel In partlabels
                                        If plabel.LabelType = ppcbLabelTypeRefDesignator Then
                                            If plabel.HorzJustification <> mplabel.HorzJustification Then plabel.HorzJustification = mplabel.HorzJustification
                                            If plabel.VertJustification <> mplabel.VertJustification Then plabel.VertJustification = mplabel.VertJustification
                                            tempX = mplabel.PositionX(ppcbUnits) - mpart.PositionX(ppcbUnits) + offsetX
                                            tempY = mplabel.PositionY(ppcbUnits) - mpart.PositionY(ppcbUnits) + offsetY
                                            If plabel.PositionX(ppcbUnits) <> tempX Then plabel.PositionX(ppcbUnits) = tempX
                                            If plabel.PositionY(ppcbUnits) <> tempY Then plabel.PositionY(ppcbUnits) = tempY
                                            If plabel.Orientation <> mplabel.Orientation Then plabel.Orientation = mplabel.Orientation
                                            If plabel.layer <> mplabel.layer Then plabel.layer = mplabel.layer
                                            If plabel.LineWidth <> mplabel.LineWidth Then plabel.LineWidth = mplabel.LineWidth
                                            If plabel.Height <> mplabel.Height Then plabel.Height = mplabel.Height
                                            Exit For
                                        End If
                                    Next plabel
                                End If 
                                                              
                                Exit For
	                        End If
                        Next mpart
                    End If        ' skip
                End If            ' req attrib exist
	        End If                ' not master part
	    Next part
	End If                        ' not enough masterparts
	Debug.Print Now
	
	If blnQuestionableData Then
	    MsgAnswer = MsgBox("Questionable Data found", vbYes)
	End If
	Rem
End Sub

Sub CalcOffsets(block As String,cell As String)
    Dim temp As Long, temp2 As Long
    
    offsetX = 3000   'default offset is nothing good calculated
    offsetY = 2000
    If ppcbUnits = ppcbUnitMetric Then
         offsetX = 600   'default offset is nothing good calculated
         offsetY = 400
    End If
    
    ' hex value attribute will be prefixed with "&H"
    'If InStr(1,"&H",cell) > 0 Then
    temp2 = Val (cell)
    If temp2 >= 0 And temp2 < 16 And block >= "A" And block <= "D" Then
        offsetX = GridData(temp2)(0)
        offsetY = GridData(temp2)(1)    
        
        temp = 0
        Select Case block
        Case "A"
        Case "B"
            temp = 1
        Case "C"
            temp = 2
        Case "D"
            temp = 3
        End Select

        offsetX = offsetX * GridSymmetry(temp)(0) + GridOffset(temp)(0)
        offsetY = offsetY * GridSymmetry(temp)(1) + GridOffset(temp)(1)
    Else
        blnQuestionableData = True
    End If
' old code
'    Select Case block
'    Case "A"
'    Case "B"
'        offsetY = -offsetY           ' B mirror of A in x axis (y=0)
'    Case "C"
'        offsetX = offsetX + Block_offsetX     ' C is x axix offset from A
'    Case "D"
'        offsetX = offsetX + Block_offsetX     ' D is offset from B
'        offsetY = -offsetY
'    End Select
    
    Debug.Print " : ";block;cell; " : ";offsetX; " ";offsetY
    
End Sub
