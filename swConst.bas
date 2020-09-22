Attribute VB_Name = "swConst"

'----
' Document Types
'----


Public Enum swDocumentTypes_e
    swDocNONE = 0                       ' Used to be TYPE_NONE
    swDocPART = 1                       ' Used to be TYPE_PART
    swDocASSEMBLY = 2           ' Used to be TYPE_ASSEMBLY
    swDocDRAWING = 3            ' Used to be TYPE_DRAWING
        swDocSDM = 4                    ' Solid data manager.
End Enum

'----
' Selection Types
'----

' The following are the possible type ids returned by the function
'     ISelectionMgr::GetSelectedObjectType.
' The string names to the right of the type id definition is the "type name"
'     used by the methods:  IModelDoc::SelectByID && AndSelectByID

Public Enum swSelectType_e

        swSelNOTHING = 0

        swSelEDGES = 1                          ' "EDGE"
        swSelFACES = 2                          ' "FACE"
        swSelVERTICES = 3               ' "VERTEX"
        swSelDATUMPLANES = 4            ' "PLANE"
        swSelDATUMAXES = 5              ' "AXIS"

        swSelDATUMPOINTS = 6            ' "DATUMPOINT"
        swSelOLEITEMS = 7               ' "OLEITEM"
        swSelATTRIBUTES = 8             ' "ATTRIBUTE"
        swSelSKETCHES = 9               ' "SKETCH"
        swSelSKETCHSEGS = 10            ' "SKETCHSEGMENT"

        swSelSKETCHPOINTS = 11          ' "SKETCHPOINT"
        swSelDRAWINGVIEWS = 12          ' "DRAWINGVIEW"
        swSelGTOLS = 13                         ' "GTOL"
        swSelDIMENSIONS = 14            ' "DIMENSION"
        swSelNOTES = 15                         ' "NOTE"

        swSelSECTIONLINES = 16          ' "SECTIONLINE"
        swSelDETAILCIRCLES = 17         ' "DETAILCIRCLE"
        swSelSECTIONTEXT = 18           ' "SECTIONTEXT"
        swSelSHEETS = 19                        ' "SHEET"
        swSelCOMPONENTS = 20            ' "COMPONENT"

        swSelMATES = 21                         ' "MATE"
        swSelBODYFEATURES = 22          ' "BODYFEATURE"
        swSelREFCURVES = 23             ' "REFCURVE"
        swSelEXTSKETCHSEGS = 24         ' "EXTSKETCHSEGMENT"
        swSelEXTSKETCHPOINTS = 25       ' "EXTSKETCHPOINT"

        swSelHELIX = 26                         ' "HELIX" (is this wrong?)
        swSelREFERENCECURVES = 26       ' "REFERENCECURVES"
        swSelREFSURFACES = 27           ' "REFSURFACE"
        swSelCENTERMARKS = 28           ' "CENTERMARKS"
        swSelINCONTEXTFEAT = 29         ' "INCONTEXTFEAT"
        swSelMATEGROUP = 30             ' "MATEGROUP"

        swSelBREAKLINES = 31            ' "BREAKLINE"
        swSelINCONTEXTFEATS = 32        ' "INCONTEXTFEATS"
        swSelMATEGROUPS = 33            ' "MATEGROUPS"
        swSelSKETCHTEXT = 34            ' "SKETCHTEXT"
        swSelSFSYMBOLS = 35             ' "SFSYMBOL"

        swSelDATUMTAGS = 36             ' "DATUMTAG"
        swSelCOMPPATTERN = 37           ' "COMPPATTERN"
        swSelWELDS = 38                         ' "WELD"
        swSelCTHREADS = 39              ' "CTHREAD"
        swSelDTMTARGS = 40              ' "DTMTARG"

        swSelPOINTREFS = 41             ' "POINTREF"
        swSelDCABINETS = 42             ' "DCABINET"
        swSelEXPLVIEWS = 43             ' "EXPLODEDVIEWS"
        swSelEXPLSTEPS = 44             ' "EXPLODESTEPS"
        swSelEXPLLINES = 45             ' "EXPLODELINES"

        swSelSILHOUETTES = 46           ' "SILHOUETTE"
        swSelCONFIGURATIONS = 47        ' "CONFIGURATIONS"
        swSelOBJHANDLES = 48
        swSelARROWS = 49                        ' "VIEWARROW"
        swSelZONES = 50                         ' "ZONES"

        swSelREFEDGES = 51              ' "REFERENCE-EDGE"
        swSelREFFACES = 52
        swSelREFSILHOUETTE = 53
        swSelBOMS = 54                          ' "BOM"
        swSelEQNFOLDER = 55             ' "EQNFOLDER"

        swSelSKETCHHATCH = 56           ' "SKETCHHATCH"
        swSelIMPORTFOLDER = 57          ' "IMPORTFOLDER"
        swSelVIEWERHYPERLINK = 58       ' "HYPERLINK"
        swSelMIDPOINTS = 59
        swSelCUSTOMSYMBOLS = 60         ' "CUSTOMSYMBOL"

        swSelCOORDSYS = 61              ' "COORDSYS"
        swSelDATUMLINES = 62            ' "REFLINE"
        swSelROUTECURVES = 63
        swSelBOMTEMPS = 64              ' "BOMTEMP"
        swSelROUTEPOINTS = 65           ' "ROUTEPOINT"

        swSelCONNECTIONPOINTS = 66      ' "CONNECTIONPOINT"
        swSelROUTESWEEPS = 67
        swSelPOSGROUP = 68              ' "POSGROUP"
        swSelBROWSERITEM = 69           ' "BROWSERITEM"
        swSelFABRICATEDROUTE = 70       ' "ROUTEFABRICATED"

        swSelSKETCHPOINTFEAT = 71       ' "SKETCHPOINTFEAT"
        swSelEMPTYSPACE = 72            ' (is this wrong?)
        swSelCOMPSDONTOVERRIDE = 72
        swSelLIGHTS = 73                        ' "LIGHTS"
        swSelWIREBODIES = 74
        swSelSURFACEBODIES = 75         ' "SURFACEBODY"

        swSelSOLIDBODIES = 76           ' "SOLIDBODY"
        swSelFRAMEPOINT = 77            ' "FRAMEPOINT"
        swSelSURFBODIESFIRST = 78
        swSelMANIPULATORS = 79          ' "MANIPULATOR"
        swSelPICTUREBODIES = 80         ' "PICTURE BODY"

        swSelSOLIDBODIESFIRST = 81

        'swSelEVERYTHING     = 4294967293
        'swSelLOCATIONS = 4294967294
        'swSelUNSUPPORTED       = 4294967295
End Enum

'----
' Events Notifications
'----

Public Enum swViewNotify_e   ' For IModelView ( DIID_DSldWorksEvents

        swViewRepaintNotify = 1
        swViewChangeNotify = 2
        swViewDestroyNotify = 3
        swViewRepaintPostNotify = 4
        swViewBufferSwapNotify = 5
        swViewDestroyNotify2 = 6
End Enum

Public Enum swFMViewNotify_e   ' For IFeatMgrView ( DIID_DSldWorksEvents

        swFMViewActivateNotify = 1
        swFMViewDeactivateNotify = 2
        swFMViewDestroyNotify = 3
End Enum

Public Enum swPartNotify_e    ' For IPartDoc ( DIID_DPartDocEvents )

        swPartRegenNotify = 1
        swPartDestroyNotify = 2
        swPartRegenPostNotify = 3
        swPartViewNewNotify = 4
        swPartNewSelectionNotify = 5
        swPartFileSaveNotify = 6
        swPartFileSaveAsNotify = 7
        swPartLoadFromStorageNotify = 8
        swPartSaveToStorageNotify = 9
        swPartConfigChangeNotify = 10
        swPartConfigChangePostNotify = 11
        swPartAutoSaveNotify = 12
        swPartAutoSaveToStorageNotify = 13
        swPartViewNewNotify2 = 14
        swPartLightingDialogCreateNotify = 15
        swPartAddItemNotify = 16
        swPartRenameItemNotify = 17
        swPartDeleteItemNotify = 18
        swPartModifyNotify = 19
        swPartFileReloadNotify = 20
        swPartAddCustomPropertyNotify = 21
        swPartChangeCustomPropertyNotify = 22
        swPartDeleteCustomPropertyNotify = 23
        swPartFeatureEditPreNotify = 24
        swPartFeatureSketchEditPreNotify = 25
        swPartFileSaveAsNotify2 = 26
        swPartDeleteSelectionPreNotify = 27
        swPartFileReloadPreNotify = 28
        swPartBodyVisibleChangeNotify = 29
End Enum

Public Enum swDrawingNotify_e    ' For IDrawingDoc ( DIID_DDrawingDocEvents )

        swDrawingRegenNotify = 1
        swDrawingDestroyNotify = 2
        swDrawingRegenPostNotify = 3
        swDrawingViewNewNotify = 4
        swDrawingNewSelectionNotify = 5
        swDrawingFileSaveNotify = 6
        swDrawingFileSaveAsNotify = 7
        swDrawingLoadFromStorageNotify = 8
        swDrawingSaveToStorageNotify = 9
        swDrawingAutoSaveNotify = 10
        swDrawingAutoSaveToStorageNotify = 11
        swDrawingConfigChangeNotify = 12
        swDrawingConfigChangePostNotify = 13
        swDrawingViewNewNotify2 = 14
        swDrawingAddItemNotify = 15
        swDrawingRenameItemNotify = 16
        swDrawingDeleteItemNotify = 17
        swDrawingModifyNotify = 18
        swDrawingFileReloadNotify = 19
        swDrawingAddCustomPropertyNotify = 20
        swDrawingChangeCustomPropertyNotify = 21
        swDrawingDeleteCustomPropertyNotify = 22
        swDrawingFileSaveAsNotify2 = 23
        swDrawingDeleteSelectionPreNotify = 24
        swDrawingFileReloadPreNotify = 25
End Enum

Public Enum swAssemblyNotify_e ' For IAssemblyDoc ( DIID_DAssemblyDocEvents )

        swAssemblyRegenNotify = 1
        swAssemblyDestroyNotify = 2
        swAssemblyRegenPostNotify = 3
        swAssemblyViewNewNotify = 4
        swAssemblyNewSelectionNotify = 5
        swAssemblyFileSaveNotify = 6
        swAssemblyFileSaveAsNotify = 7
        swAssemblyLoadFromStorageNotify = 8
        swAssemblySaveToStorageNotify = 9
        swAssemblyConfigChangeNotify = 10
        swAssemblyConfigChangePostNotify = 11
        swAssemblyAutoSaveNotify = 12
        swAssemblyAutoSaveToStorageNotify = 13
        swAssemblyBeginInContextEditNotify = 14
        swAssemblyEndInContextEditNotify = 15
        swAssemblyViewNewNotify2 = 16
        swAssemblyLightingDialogCreateNotify = 17
        swAssemblyAddItemNotify = 18
        swAssemblyRenameItemNotify = 19
        swAssemblyDeleteItemNotify = 20
        swAssemblyModifyNotify = 21
        swAssemblyComponentStateChangeNotify = 22
        swAssemblyFileDropNotify = 23
        swAssemblyFileReloadNotify = 24
        swAssemblyComponentStateChangeNotify2 = 25
        swAssemblyAddCustomPropertyNotify = 26
        swAssemblyChangeCustomPropertyNotify = 27
        swAssemblyDeleteCustomPropertyNotify = 28
        swAssemblyFeatureEditPreNotify = 29
        swAssemblyFeatureSketchEditPreNotify = 30
        swAssemblyFileSaveAsNotify2 = 31
        swAssemblyInterferenceNotify = 32
        swAssemblyDeleteSelectionPreNotify = 33
        swAssemblyFileReloadPreNotify = 34
        swAssemblyComponentMoveNotify = 35
        swAssemblyComponentVisibleChangeNotify = 36
        swAssemblyBodyVisibleChangeNotify = 37

End Enum

Public Enum swAppNotify_e    ' For ISldWorks ( DIID_DSldWorksEvents )

        swAppFileOpenNotify = 1
        swAppFileNewNotify = 2
        swAppDestroyNotify = 3
        swAppActiveDocChangeNotify = 4
        swAppActiveModelDocChangeNotify = 5
        swAppPropertySheetCreateNotify = 6
        swAppNonNativeFileOpenNotify = 7
        swAppLightSheetCreateNotify = 8
        swAppDocumentConversionNotify = 9
        swAppLightweightComponentOpenNotify = 10
        swAppDocumentLoadNotify = 11
        swAppFileNewNotify2 = 12
        swAppFileOpenNotify2 = 13
        swAppReferenceNotFoundNotify = 14
        swAppPromptForFilenameNotify = 15
End Enum

Public Enum swPropertySheetNotify_e

        swPropertySheetDestroyNotify = 1
        swPropertySheetHelpNotify = 2
End Enum

'----
' Parameter Types
'----
Public Enum swParamType_e    ' For use with IAttributeDef::AddParameter (for example)

        swParamTypeDouble = 0
        swParamTypeString = 1
        swParamTypeInteger = 2
        swParamTypeDVector = 3
End Enum

'----
' The following is for angular dimension info returned GetDimensionInfo()
'----
Public Enum swQuadant_e

        swQuadUnknown = 0
        swQuadPosQ1 = 1
        swQuadNegQ1 = 2
        swQuadPosQ2 = 3
        swQuadNegQ2 = 4
End Enum

'----
' The following Public Enum is for interpreting ellipse data
'----
Public Enum swEllipsePts_e

        swEllipseStartPt = 0
        swEllipseEndPt = 1
        swEllipseCenterPt = 2
        swEllipseMajorPt = 3
        swEllipseMinorPt = 4
End Enum

Public Enum swParabolaPts_e

        swParabolaStartPt = 0
        swParabolaEndPt = 1
        swParabolaFocusPt = 2
        swParabolaApexPt = 3
End Enum

'----
' The following define gtol symbol indices
'----
Public Enum swGtolMatCondition_e

        swMcNONE = 0
        swMcMMC = 1
        swMcRFS = 2
        swMcLMC = 3

        swMsNONE = 4
        swMsPROJTOLZONE = 5
        swMsDIA = 6
        swMsSPHDIA = 7
        swMsRAD = 8
        swMsSPHRAD = 9
        swMsREF = 10
        swMsARCLEN = 11
End Enum

Public Enum swGtolGeomCharSymbol_e

        swGcsNONE = 12
        swGcsSYMMETRY = 13
        swGcsSTRAIGHT = 14
        swGcsFLAT = 15
        swGcsROUND = 16
        swGcsCYL = 17
        swGcsLINEPROF = 18
        swGcsSURFPROF = 19
        swGcsANG = 20
        swGcsPERP = 21
        swGcsPARALLEL = 22
        swGcsPOSITION = 23
        swGcsCONC = 24
        swGcsCIRCRUNOUT = 25
        swGcsTOTALRUNOUT = 26
        swGcsCIRCOPENRUNOUT = 27
        swGcsTOTALOPENRUNOUT = 28
End Enum

Public Enum swMateType_e

        swMateCOINCIDENT = 0
        swMateCONCENTRIC = 1
        swMatePERPENDICULAR = 2
        swMatePARALLEL = 3
        swMateTANGENT = 4
        swMateDISTANCE = 5
        swMateANGLE = 6
        swMateUNKNOWN = 7
        swMateSYMMETRIC = 8
        swMateCAMFOLLOWER = 9
End Enum

' Enumerations for Detail View Creation
Public Enum swDetCircleShowType_e

        swDetCirclePROFILE = 0
        swDetCircleCIRCLE = 1
        swDetCircleDONTSHOW = 2
End Enum

' Enumerations for Detail View Style
Public Enum swDetViewStyle_e

        swDetViewSTANDARD = 0
        swDetViewBROKEN = 1
        swDetViewLEADER = 2
        swDetViewNOLEADER = 3
        swDetViewCONNECTED = 4
End Enum

' This Public Enum has been changed to correct improper mate alignment mapping
Public Enum swMateAlign_e

        ' Use the three corrected enum's below
        swMateAlignALIGNED = 0
        swMateAlignANTI_ALIGNED = 1
        swMateAlignCLOSEST = 2

        ' Old incorrect enum's retained for backwards compatability
        ' Avoid using the three incorrect enum's below if possible
        swAlignNONE = 0
        swAlignSAME = 1
        swAlignAGAINST = 2
End Enum


Public Enum swDisplayMode_e

        swWIREFRAME = 0
        swHIDDEN_GREYED = 1
        swHIDDEN = 2
End Enum

Public Enum swArrowStyle_e

        swOPEN_ARROWHEAD = 0
        swCLOSED_ARROWHEAD = 1
        swSLASH_ARROWHEAD = 2
        swDOT_ARROWHEAD = 3
        swORIGIN_ARROWHEAD = 4
        swWIDE_ARROWHEAD = 5
        swISOWIDE_ARROWHEAD = 6
        swRUS_ARROWHEAD = 7
        swCLOSETOP_ARROWHEAD = 8
        swCLOSEBOT_ARROWHEAD = 9
        swNO_ARROWHEAD = 10
End Enum

Public Enum swLeaderSide_e

        swLS_SMART = 0
        swLS_LEFT = 1
        swLS_RIGHT = 2
End Enum


'----
' The following define Surface Finish Symbol types and options
' Used by InsertSurfaceFinishSymbol ModifySurfaceFinishSymbol
'----
Public Enum swSFSymType_e
 
        swSFBasic = 0
        swSFMachining_Req = 1
        swSFDont_Machine = 2
        swSFJIS_Surface_Texture_1 = 3                   ' Add next 5 JIS types, 08/26/99
        swSFJIS_Surface_Texture_2 = 4
        swSFJIS_Surface_Texture_3 = 5
        swSFJIS_Surface_Texture_4 = 6
        swSFJIS_No_Machining = 7
End Enum

Public Enum swSFLaySym_e
 
        swSFNone = 0
        swSFCircular = 1
        swSFCross = 2
        swSFMultiDir = 3
        swSFParallel = 4
        swSFPerp = 5
        swSFRadial = 6
        swSFParticulate = 7
End Enum

' The different possibilities for types of texts in a Surface Finish symbol. (SFSymbol::Get/SetText)
Public Enum swSurfaceFinishSymbolText_e

        swSFSymbolMaterialRemovalAllowance = 1
        swSFSymbolProductionMethod = 2
        swSFSymbolSamplingLength = 3
        swSFSymbolOtherRoughnessValue = 4
        swSFSymbolMaximumRoughness = 5
        swSFSymbolMinimumRoughness = 6
        swSFSymbolRoughnessSpacing = 7
End Enum

Public Enum swLeaderStyle_e
 
        swNO_LEADER = 0
        swSTRAIGHT = 1
        swBENT = 2
End Enum

'----
' Balloon Information.  swBS_SplitCirc is not valid for Notes only for Balloons
'----
Public Enum swBalloonStyle_e

        swBS_None = 0
        swBS_Circular = 1
        swBS_Triangle = 2
        swBS_Hexagon = 3
        swBS_Box = 4
        swBS_Diamond = 5
        swBS_SplitCirc = 6
        swBS_Pentagon = 7
        swBS_FlagPentagon = 8
        swBS_FlagTriangle = 9
End Enum

Public Enum swBalloonFit_e

        swBF_Tightest = 0
        swBF_1Char = 1
        swBF_2Chars = 2
        swBF_3Chars = 3
        swBF_4Chars = 4
        swBF_5Chars = 5
End Enum

'----
' The following define length and angle unit types
'----
Public Enum swLengthUnit_e

        swMM = 0
        swCM = 1
        swMETER = 2
        swINCHES = 3
        swFEET = 4
        swFEETINCHES = 5
End Enum

Public Enum swAngleUnit_e

        swDEGREES = 0
        swDEG_MIN = 1
        swDEG_MIN_SEC = 2
        swRADIANS = 3
End Enum

Public Enum swFractionDisplay_e

        swNONE = 0
        swDECIMAL = 1
        swFRACTION = 2
End Enum

'----
' Drawing Paper Sizes
'----
Public Enum swDwgPaperSizes_e

        swDwgPaperAsize = 0
        swDwgPaperAsizeVertical = 1
        swDwgPaperBsize = 2
        swDwgPaperCsize = 3
        swDwgPaperDsize = 4
        swDwgPaperEsize = 5
        swDwgPaperA4size = 6
        swDwgPaperA4sizeVertical = 7
        swDwgPaperA3size = 8
        swDwgPaperA2size = 9
        swDwgPaperA1size = 10
        swDwgPaperA0size = 11
        swDwgPapersUserDefined = 12
End Enum


'----
' Drawing Templates
'----
Public Enum swDwgTemplates_e

        swDwgTemplateAsize = 0
    swDwgTemplateAsizeVertical = 1
        swDwgTemplateBsize = 2
        swDwgTemplateCsize = 3
        swDwgTemplateDsize = 4
        swDwgTemplateEsize = 5
        swDwgTemplateA4size = 6
        swDwgTemplateA4sizeVertical = 7
        swDwgTemplateA3size = 8
        swDwgTemplateA2size = 9
        swDwgTemplateA1size = 10
        swDwgTemplateA0size = 11
        swDwgTemplateCustom = 12
        swDwgTemplateNone = 13
End Enum

'----
' Drawing Templates
'----
Public Enum swStandardViews_e

        swFrontView = 1
        swBackView = 2
        swLeftView = 3
        swRightView = 4
        swTopView = 5
        swBottomView = 6
        swIsometricView = 7
        swTrimetricView = 8
        swDimetricView = 9
End Enum

'----
' Repaint Notification types
'----
Public Enum swRepaintTypes_e

        swStandardUpdate = 0
        swLightUpdate = 1
        swMaterialUpdate = 2
        swSectionedUpdate = 3
        swExplodedUpdate = 4
        swInsertSketchUpdate = 5
        swViewDisplayUpdate = 6
        swDamageRepairUpdate = 7
        swSelectionUpdate = 8
        swSectionedExitUpdate = 9
        swScrollViewUpdate = 10
End Enum

'----
' User Interface State
'----
Public Enum swUIStates_e

        swIsHiddenInFeatureMgr = 1
End Enum

'----
' Type names
'----

' Body Features
Public Const swTnChamfer As String = "Chamfer"
Public Const swTnFillet As String = "Fillet"
Public Const swTnCavity As String = "Cavity"
Public Const swTnDraft As String = "Draft"
Public Const swTnMirrorSolid As String = "MirrorSolid"
Public Const swTnCirPattern As String = "CirPattern"
Public Const swTnLPattern As String = "LPattern"
Public Const swTnMirrorPattern As String = "MirrorPattern"
Public Const swTnShell As String = "Shell"
Public Const swTnBlend As String = "Blend"
Public Const swTnBlendCut As String = "BlendCut"
Public Const swTnExtrusion As String = "Extrusion"
Public Const swTnBoss As String = "Boss"
Public Const swTnCut As String = "Cut"
Public Const swTnRefCurve As String = "RefCurve"
Public Const swTnRevolution As String = "Revolution"
Public Const swTnRevCut As String = "RevCut"
Public Const swTnSweep As String = "Sweep"
Public Const swTnSweepCut As String = "SweepCut"
Public Const swTnStock As String = "Stock"
Public Const swTnSurfCut As String = "SurfCut"
Public Const swTnThicken As String = "Thicken"
Public Const swTnThickenCut As String = "ThickenCut"
Public Const swTnVarFillet As String = "VarFillet"
Public Const swTnSketchHole As String = "SketchHole"
Public Const swTnHoleWzd As String = "HoleWzd"
Public Const swTnImported As String = "Imported"
Public Const swTnBaseBody As String = "BaseBody"
Public Const swTnDerivedLPattern As String = "DerivedLPattern"
Public Const swTnCosmeticThread As String = "CosmeticThread"

' Sheet Metal features
Public Const swTnSheetMetal As String = "SheetMetal"
Public Const swTnFlattenBends As String = "FlattenBends"
Public Const swTnProcessBends As String = "ProcessBends"
Public Const swTnOneBend As String = "OneBend"
Public Const swTnBaseFlange As String = "SMBaseFlange"
Public Const swTnSketchBend As String = "SketchBend"
Public Const swTnSM3dBend As String = "SM3dBend"
Public Const swTnEdgeFlange As String = "EdgeFlange"
Public Const swTnFlatPattern As String = "FlatPattern"

' Drawing Related
Public Const swTnCenterMark As String = "CenterMark"
Public Const swTnDrSheet As String = "DrSheet"
Public Const swTnAbsoluteView As String = "AbsoluteView"
Public Const swTnDetailView As String = "DetailView"
Public Const swTnRelativeView As String = "RelativeView"
Public Const swTnSectionPartView As String = "SectionPartView"
Public Const swTnSectionAssemView As String = "SectionAssemView"
Public Const swTnUnfoldedView As String = "UnfoldedView"
Public Const swTnAuxiliaryView As String = "AuxiliaryView"
Public Const swTnDetailCircle As String = "DetailCircle"
Public Const swTnDrSectionLine As String = "DrSectionLine"

' Assembly Related
Public Const swTnMateCoincident As String = "MateCoincident"
Public Const swTnMateConcentric As String = "MateConcentric"
Public Const swTnMateDistanceDim As String = "MateDistanceDim"
Public Const swTnMateParallel As String = "MateParallel"
Public Const swTnMateTangent As String = "MateTangent"
Public Const swTnReference As String = "Reference"

' Reference Geometry
Public Const swTnRefPlane As String = "RefPlane"
Public Const swTnRefAxis As String = "RefAxis"
Public Const swTnReferenceCurve As String = "ReferenceCurve"
Public Const swTnRefSurface As String = "RefSurface"
Public Const swTnCoordinateSystem As String = "CoordSys"

' Misc
Public Const swTnAttribute As String = "Attribute"
Public Const swTnProfileFeature As String = "ProfileFeature"

' Symbol markers
Public Const SYMBOL_MARKER_START As String = "<"
Public Const SYMBOL_MARKER_END As String = ">"
Public Const SYMBOL_MARKER_SPACE As String = "-"


'----
' Surface Types.  For use with Surface::Identity method.
'----
Public Const PLANE_TYPE As Integer = 4001
Public Const CYLINDER_TYPE As Integer = 4002
Public Const CONE_TYPE As Integer = 4003
Public Const SPHERE_TYPE As Integer = 4004
Public Const TORUS_TYPE As Integer = 4005
Public Const BSURF_TYPE As Integer = 4006
Public Const BLEND_TYPE As Integer = 4007
Public Const OFFSET_TYPE As Integer = 4008
Public Const EXTRU_TYPE As Integer = 4009
Public Const SREV_TYPE As Integer = 4010

'----
' Curve Types.  For use with Curve::Identity method.
'----
Public Const LINE_TYPE As Integer = 3001
Public Const CIRCLE_TYPE As Integer = 3002
Public Const ELLIPSE_TYPE As Integer = 3003
Public Const INTERSECTION_TYPE As Integer = 3004
Public Const BCURVE_TYPE As Integer = 3005
Public Const SPCURVE_TYPE As Integer = 3006
Public Const CONSTPARAM_TYPE As Integer = 3008
Public Const TRIMMED_TYPE As Integer = 3009

'----
' This is the beginning of time. Used to initialize su_CTime.
'----
Public Const TIME_ORIGIN As String = "1990 1, 1, 0, 0, 0"



' Items that can be configured to have a line style in drawings.
Public Enum swLineTypes_e

        swLF_VISIBLE = 0
        swLF_HIDDEN = 1
        swLF_SKETCH = 2
        swLF_DETAIL = 3
        swLF_SECTION = 4
        swLF_DIMENSION = 5
        swLF_CENTER = 6
        swLF_HATCH = 7
        swLF_TANGENT = 8
End Enum

' Dimension tolerance types
Public Enum swTolType_e

        swTolNONE = 0
        swTolBASIC = 1
        swTolBILAT = 2
        swTolLIMIT = 3
        swTolSYMMETRIC = 4
        swTolMIN = 5
        swTolMAX = 6
        swTolMETRIC = 7
End Enum

' Tolerances which the user can set using Modeler::SetTolerances
Public Enum swTolerances_e

        swBSCurveOutputTol = 0                  '3D bspline curve output tolerance (meters)
        swBSCurveNonRationalOutputTol = 1       '3D non-rational bspline curve output tolerance (meters)
        swUVCurveOutputTol = 2                          '2D trim curve output tolerance (fraction of characteristic min. face dimension)
        swSurfChordTessellationTol = 3          'chord tolerance or deviation for tessellation for surfaces
        swSurfAngularTessellationTol = 4        'angular tolerance or deviation for tessellation for surfaces
        swCurveChordTessellationTol = 5         'chord tolerance or deviation for tessellation for curves
End Enum

'----
' Mate Entity Types
'
'  The following are the possible mate entity type ids returned by the function
'  IMateEntity::GetEntityType.
'----
Public Enum swMateEntityTypes_e

        swMateUnsupported = 0
        swMatePoint = 1
        swMateLine = 2
        swMatePlane = 3
        swMateCylinder = 4
        swMateCone = 5
End Enum

'----
' Attribute Callback Support
'
'  The following are the possible callback types for IAttributeDefs
'----
Public Enum swAttributeCallbackTypes_e

        swACBDelete = 0
End Enum

Public Enum swAttributeCallbackOptions_e

        swACBRequiresCallback = 1
End Enum

Public Enum swAttributeCallbackReturnValues_e

        swACBDeleteIt = 1
End Enum

' Text reference point position
Public Enum swTextPosition_e

        swUPPER_LEFT = 0
        swLOWER_LEFT = 1
        swCENTER = 2
        swUPPER_RIGHT = 3
        swLOWER_RIGHT = 4
        swUPPER_CENTER = 5
End Enum

'----
' The following are the different types of topology resulting from a call to GetTrimCurves
'----
Public Enum swTopologyTypes_e

        swTopologyNull = 0
        swTopologyCoEdge = 1
        swTopologyVertex = 2
End Enum


'----
' Attributes associated entity state
'----
Public Enum swAssociatedEntityStates_e

        swIsEntityInvalid = 0
        swIsEntitySuppressed = 1
        swIsEntityAmbiguous = 2
        swIsEntityDeleted = 3
End Enum

'---
' Search Folder Types
'---
Public Enum swSearchFolderTypes_e

        swDocumentType = 0
End Enum


'---
' User Preference Toggles.
' The different User Preference Toggles for GetUserPreferenceToggle & SetUserPreferenceToggle
'---
Public Enum swUserPreferenceToggle_e

        swUseFolderSearchRules = 0
        swDisplayArcCenterPoints = 1
        swDisplayEntityPoints = 2
        swIgnoreFeatureColors = 3
        swDisplayAxes = 4
        swDisplayPlanes = 5
        swDisplayOrigins = 6
        swDisplayTemporaryAxes = 7
        swDxfMapping = 8
        swSketchAutomaticRelations = 9
        swInputDimValOnCreate = 10
        swFullyConstrainedSketchMode = 11
        swXTAssemSaveFormat = 12
        swDisplayCoordSystems = 13
        swExtRefOpenReadOnly = 14
        swExtRefNoPromptOrSave = 15
        swExtRefMultipleContexts = 16
        swExtRefAutoGenNames = 17
        swExtRefUpdateCompNames = 18
        swDisplayReferencePoints = 19
        swUseShadedFaceHighlight = 20
        swDXFDontShowMap = 21
        swThumbnailGraphics = 22
        swUseAlphaTransparency = 23
        swDynamicDrawingViewActivation = 24
        swAutoLoadPartsLightweight = 25
        swIGESStandardSetting = 26
        swIGESNurbsSetting = 27
        swTiffPrintScaleToFit = 28
        swDisplayVirtualSharps = 29
        swUpdateMassPropsDuringSave = 30
        swDisplayAnnotations = 31
        swDisplayFeatureDimensions = 32
        swDisplayReferenceDimensions = 33
        swDisplayAnnotationsUseAssemblySettings = 34
        swDisplayNotes = 35
        swDisplayGeometricTolerances = 36
        swDisplaySurfaceFinishSymbols = 37
        swDisplayWeldSymbols = 38
        swDisplayDatums = 39
        swDisplayDatumTargets = 40
        swDisplayCosmeticThreads = 41

        swDetailingDisplayWithBrokenLeaders = 42
        swDetailingDualDimensions = 43
        swDetailingDisplayDatumsPer1982 = 44
        swDetailingDisplayAlternateSection = 45
        swDetailingCenterMarkShowLines = 46
        swDetailingFixedSizeWeldSymbol = 47
        swDetailingDimsShowParenthesisByDefault = 48
        swDetailingDimsSnapTextToGrid = 49
        swDetailingDimsCenterText = 50
        swDetailingRadialDimsDisplay2ndOutsideArrow = 51
        swDetailingRadialDimsArrowsFollowText = 52
        swDetailingDimLeaderOverrideStandard = 53
        swDetailingNotesDisplayWithBentLeader = 54
        swDisplayTextAtSameSizeAlways = 55
        swDisplayOnlyInViewOfCreation = 56
        swGridDisplay = 57
        swGridDisplayDashed = 58
        swGridAutomaticScaling = 59
        swSnapToPoints = 60
        swSnapToAngle = 61
        swUnitsLinearRoundToNearestFraction = 62
        swUnitsLinearFeetAndInchesFormat = 63

        swFeatureManagerEnsureVisible = 64
        swFeatureManagerNameFeatureWhenCreated = 65
        swFeatureManagerKeyboardNavigation = 66
        swFeatureManagerDynamicHighlight = 67
        swColorsGradientPartBackground = 68

        swSTLBinaryFormat = 69
        swSTLShowInfoOnSave = 70
        swSTLDontTranslateToPositive = 71
        swSTLComponentsIntoOneFile = 72
        swSTLCheckForInterference = 73

        swOpenLastUsedDocumentAtStart = 74
        swSingleCommandPerPick = 75
        swShowDimensionNames = 76
        swShowErrorsEveryRebuild = 77
        swMaximizeDocumentOnOpen = 78
        swEditDesignTableInSeparateWindow = 80
        swEnablePropertyManager = 81
        swUseSystemSeparatorForDims = 82
        swUseEnglishLanguage = 83
        swDrawingAutomaticModelDimPlacement = 84
        swDrawingDisplayViewBorders = 85
        swAutomaticScaling3ViewDrawings = 86
        swDrawingAutomaticBomUpdate = 87
        swDrawingSelectHiddenEntities = 88
        swDrawingCreateDetailAsCircle = 89
        swAutomaticDrawingViewUpdate = 90
        swDrawingDetailInferCorner = 91
        swDrawingDetailInferCenter = 92
        swDrawingViewShowContentsWhileDragging = 93
        swSketchAlternateSplineCreation = 94
        swSketchInferFromModel = 95
        swSketchPromptToCloseSketch = 96
        swSketchCreateSketchOnNewPart = 97
        swSketchOverrideDimensionsOnDrag = 98
        swSketchDisplayPlaneWhenShaded = 99
        swSketchOverdefiningDimsPromptToSetState = 100
        swSketchOverdefiningDimsSetDrivenByDefault = 101
        swPerformanceVerifyOnRebuild = 102
        swPerformanceDynamicUpdateOnMove = 103
        swPerformanceAlwaysGenerateCurvature = 104
        swPerformanceWin95ZoomClipping = 105

        swIGESDuplicateEntities = 106
        swIGESHighTrimCurveAccuracy = 107
        swIGESExportSketchEntities = 108
        swIGESComponentsIntoOneFile = 109
        swIGESFlattenAssemHierarchy = 110

        swAlwaysUseDefaultTemplates = 111

        swUseSimpleOpenGL = 112
        swShowRefGeomName = 113
        swUseShadedPreview = 114

        swEdgesHiddenEdgeSelectionInWireframe = 115
        swEdgesHiddenEdgeSelectionInHLR = 116
        swEdgesRepaintAfterSelectionInHLR = 117
        swEdgesHighlightFeatureEdges = 118
        swEdgesDynamicHighlight = 119
        swEdgesHighQualityDisplay = 120
        swEdgesOpenEdgesDifferentColor = 121

        swEnableConfirmationCorner = 122
        swAutoShowPropertyManager = 123
        swIncontextFeatureHolderVisibility = 124

        swTransparencyHighQualityDynamic = 125
        
        swPageSetupPrinterUsePrinterMargin = 128
        swPageSetupPrinterDrawingScaleToFit = 129
        swPageSetupPrinterPartAsmPrintWindow = 130

End Enum

'---
' User Preference Integer Values
' The different User Preference Integer Values for GetUserPreferenceIntegerValue & SetUserPreferenceIntegerValue
'---
Public Enum swUserPreferenceIntegerValue_e

        swDxfVersion = 0
        swDxfOutputFonts = 1
        swDxfMappingFileIndex = 2
        swAutoSaveInterval = 3
        swResolveLightweight = 4
        swAcisOutputVersion = 5
        swTiffScreenOrPrintCapture = 6
        swTiffImageType = 7
        swTiffCompressionScheme = 8
        swTiffPrintDPI = 9
        swTiffPrintPaperSize = 10
        swTiffPrintScaleFactor = 11
        swCreateBodyFromSurfacesOption = 12     ' Used by API CreateBodyFromSurfaces

        swDetailingDimensionStandard = 13
        swDetailingDualDimPosition = 14
        swDetailingDimTrailingZero = 15
        swDetailingArrowStyleForDimensions = 16
        swDetailingDimensionArrowPosition = 17
        swDetailingLinearDimLeaderStyle = 18
        swDetailingRadialDimLeaderStyle = 19
        swDetailingAngularDimLeaderStyle = 20
        swDetailingLinearToleranceStyle = 21
        swDetailingAngularToleranceStyle = 22
        swDetailingToleranceTextSizing = 23
        swDetailingLinearDimPrecision = 24
        swDetailingLinearTolPrecision = 25
        swDetailingAltLinearDimPrecision = 26
        swDetailingAltLinearTolPrecision = 27
        swDetailingAngularDimPrecision = 28
        swDetailingAngularTolPrecision = 29
        swDetailingNoteTextAlignment = 30
        swDetailingNoteLeaderSide = 31
        swDetailingBalloonStyle = 32
        swDetailingBalloonFit = 33
        swDetailingBOMBalloonStyle = 34
        swDetailingBOMBalloonFit = 35
        swDetailingBOMUpperText = 36
        swDetailingBOMLowerText = 37
        swDetailingArrowStyleForEdgeVertexAttachment = 38
        swDetailingArrowStyleForFaceAttachment = 39
        swDetailingArrowStyleForUnattached = 40
        swDetailingVirtualSharpStyle = 41
        swGridMinorLinesPerMajor = 42
        swSnapPointsPerMinor = 43
        swImageQualityShaded = 44
        swImageQualityWireframe = 45
        swImageQualityWireframeValue = 46
        swUnitsLinear = 47
        swUnitsLinearDecimalDisplay = 48
        swUnitsLinearDecimalPlaces = 49
        swUnitsLinearFractionDenominator = 50
        swUnitsAngular = 51
        swUnitsAngularDecimalPlaces = 52
        swLineFontVisibleEdgesThickness = 53
        swLineFontVisibleEdgesStyle = 54
        swLineFontHiddenEdgesThickness = 55
        swLineFontHiddenEdgesStyle = 56
        swLineFontSketchCurvesThickness = 57
        swLineFontSketchCurvesStyle = 58
        swLineFontDetailCircleThickness = 59
        swLineFontDetailCircleStyle = 60
        swLineFontSectionLineThickness = 61
        swLineFontSectionLineStyle = 62
        swLineFontDimensionsThickness = 63
        swLineFontDimensionsStyle = 64
        swLineFontConstructionCurvesThickness = 65
        swLineFontConstructionCurvesStyle = 66
        swLineFontCrosshatchThickness = 67
        swLineFontCrosshatchStyle = 68
        swLineFontTangentEdgesThickness = 69
        swLineFontTangentEdgesStyle = 70
        swLineFontDetailBorderThickness = 71
        swLineFontDetailBorderStyle = 72
        swLineFontCosmeticThreadThickness = 73
        swLineFontCosmeticThreadStyle = 74

        swStepAP = 75

        swHiddenEdgeDisplayDefault = 76
        swTangentEdgeDisplayDefault = 77

        swSTLQuality = 78

        swDrawingProjectionType = 79
        swDrawingPrintCrosshatchOutOfDateViews = 80
        swPerformanceAssemRebuildOnLoad = 81
        swLoadExternalReferences = 82

        swIGESRepresentation = 83
        swIGESSystem = 84
        swIGESCurveRepresentation = 85

        swViewRotationMouseSpeed = 86
        swBackupCopiesPerDocument = 87
        swCheckForOutOfDateLightweightComponents = 88

        swParasolidOutputVersion = 89

        swLineFontHideTangentEdgeThickness = 90
        swLineFontHideTangentEdgeStyle = 91
        swLineFontViewArrowThickness = 92
        swLineFontViewArrowStyle = 93

        swEdgesHiddenEdgeDisplay = 94
        swEdgesTangentEdgeDisplay = 95
        swEdgesShadedModeDisplay = 96

        swDetailingBOMStackedBalloonStyle = 97
        swDetailingBOMStackedBalloonFit = 98

        swSystemColorsViewportBackground = 99
        swSystemColorsTopGradientColor = 100
        swSystemColorsBottomGradientColor = 101
        swSystemColorsDynamicHighlight = 102
        swSystemColorsHighlight = 103
        swSystemColorsSelectedItem1 = 104
        swSystemColorsSelectedItem2 = 105
        swSystemColorsSelectedItem3 = 106
        swSystemColorsSelectedFaceShaded = 107
        swSystemColorsDrawingsVisibleModelEdge = 108
        swSystemColorsDrawingsHiddenModelEdge = 109
        swSystemColorsDrawingsPaperBorder = 110
        swSystemColorsDrawingsPaperShadow = 111
        swSystemColorsImportedDrivingAnnotation = 112
        swSystemColorsImportedDrivenAnnotation = 113
        swSystemColorsSketchOverDefined = 114
        swSystemColorsSketchFullyDefined = 115
        swSystemColorsSketchUnderDefined = 116
        swSystemColorsSketchInvalidGeometry = 117
        swSystemColorsSketchNotSolved = 118
        swSystemColorsGridLinesMinor = 119
        swSystemColorsGridLinesMajor = 120
        swSystemColorsConstructionGeometry = 121
        swSystemColorsDanglingDimension = 122
        swSystemColorsText = 123
        swSystemColorsAssemblyEditPart = 124
        swSystemColorsAssemblyEditPartHiddenLines = 125
        swSystemColorsAssemblyNonEditPart = 126
        swSystemColorsInactiveEntity = 127
        swSystemColorsTemporaryGraphics = 128
        swSystemColorsTemporaryGraphicsShaded = 129
        swSystemColorsActiveSelectionListBox = 130
        swSystemColorsSurfacesOpenEdge = 131
        swSystemColorsTreeViewBackground = 132

        swAcisOutputUnits = 133
        
        swPageSetupPrinterOrientation = 138
        swPageSetupPrinterDrawingColor = 139

End Enum

'---
' User Preference Double Values
' The different User Preference Double Values for GetUserPreferenceDoubleValue & SetUserPreferenceDoubleValue
'---
Public Enum swUserPreferenceDoubleValue_e

        swDetailingNoteFontHeight = 0
    swDetailingDimFontHeight = 1
        swSTLDeviation = 2
        swSTLAngleTolerance = 3
        swSpinBoxMetricLengthIncrement = 4
        swSpinBoxEnglishLengthIncrement = 5
        swSpinBoxAngleIncrement = 6
        swMaterialPropertyDensity = 7

' Inside SolidWorks the height and width values were switched, so that these exposed names describe
' the wrong action now.  Do not use these names anymore.
        swTiffPrintPaperWidth = 8
        swTiffPrintPaperHeight = 9
' Added these 2 names which do describe the correct action.  Use these instead.
        swTiffPrintDrawingPaperHeight = 8
        swTiffPrintDrawingPaperWidth = 9

        swDetailingCenterlineExtension = 10
        swDetailingBreakLineGap = 11
        swDetailingCenterMarkSize = 12
        swDetailingWitnessLineGap = 13
        swDetailingWitnessLineExtension = 14
        swDetailingObjectToDimOffset = 15
        swDetailingDimToDimOffset = 16
        swDetailingMaxLinearToleranceValue = 17
        swDetailingMinLinearToleranceValue = 18
        swDetailingMaxAngularToleranceValue = 19
        swDetailingMinAngularToleranceValue = 20
        swDetailingToleranceTextScale = 21
        swDetailingToleranceTextHeight = 22
        swDetailingNoteBentLeaderLength = 23
        swDetailingArrowHeight = 24
        swDetailingArrowWidth = 25
        swDetailingArrowLength = 26
        swDetailingSectionArrowHeight = 27
        swDetailingSectionArrowWidth = 28
        swDetailingSectionArrowLength = 29

        swGridMajorSpacing = 30
        swSnapToAngleValue = 31
        swImageQualityShadedDeviation = 32

        swDrawingDefaultSheetScaleNumerator = 33
        swDrawingDefaultSheetScaleDenominator = 34
        swDrawingDetailViewScale = 35
        swViewRotationArrowKeys = 36

        swMateAnimationSpeed = 37
        swViewAnimationSpeed = 38
        
        swDetailingDimBentLeaderLength = 39

        swMaterialPropertyCrosshatchScale = 40
        swMaterialPropertyCrosshatchAngle = 41
        swDrawingAreaHatchScale = 42
        swDrawingAreaHatchAngle = 43
        
        swPageSetupPrinterTopMargin = 44
        swPageSetupPrinterBottomMargin = 45
        swPageSetupPrinterLeftMargin = 46
        swPageSetupPrinterRightMargin = 47
        swPageSetupPrinterThinLineWeight = 48
        swPageSetupPrinterNormalLineWeight = 49
        swPageSetupPrinterThickLineWeight = 50
        swPageSetupPrinterThick2LineWeight = 51
        swPageSetupPrinterThick3LineWeight = 52
        swPageSetupPrinterThick4LineWeight = 53
        swPageSetupPrinterThick5LineWeight = 54
        swPageSetupPrinterThick6LineWeight = 55
        swPageSetupPrinterDrawingScale = 56
        swPageSetupPrinterPartAsmScale = 57


End Enum

'---
' User Preference String Values
' The different User Preference String Values for GetUserPreferenceStringValue & SetUserPreferenceStringValue
'---
Public Enum swUserPreferenceStringValue_e

        swFileLocationsDocuments = 1
        swFileLocationsPaletteFeatures = 2
        swFileLocationsPaletteParts = 3
        swFileLocationsPaletteFormTools = 4
        swFileLocationsBlocks = 5
        swFileLocationsDocumentTemplates = 6
        swFileLocationsSheetFormat = 7
        swDefaultTemplatePart = 8
        swDefaultTemplateAssembly = 9
        swDefaultTemplateDrawing = 10
        swBackupDirectory = 11

        swFileLocationsBendTable = 12
        
        swMaterialPropertyCrosshatchPattern = 13
        swDrawingAreaHatchPattern = 14

End Enum

'---
' User Preference String List Values
' The different User Preference String List Values for GetUserPreferenceStringListValue & SetUserPreferenceStringListValue
'---
Public Enum swUserPreferenceStringListValue_e

        swDxfMappingFiles = 0
End Enum

'---
' User Preference Text Formats
' The different User Preference Text Formats for Get/SetUserPreferenceTextFormat
'---
Public Enum swUserPreferenceTextFormat_e

        swDetailingNoteTextFormat = 0
        swDetailingDimensionTextFormat = 1
        swDetailingSectionTextFormat = 2
        swDetailingDetailTextFormat = 3
        swDetailingViewArrowTextFormat = 4
End Enum

'---
' View Display States
' The different View Display States for IModelView::GetDisplayState
'---
Public Enum swViewDisplayType_e

        swIsViewSectioned = 0
        swIsViewPerspective = 1
        swIsViewShaded = 2
        swIsViewWireFrame = 3
        swIsViewHiddenLinesRemoved = 4
        swIsViewHiddenInGrey = 5
        swIsViewCurvature = 6
End Enum

'----
' Control display of internal sketch points
'----
Public Enum swSkInternalPntOpts_e

        swSkPntsOff = 0
        swSkPntsOn = 1
        swSkPntsDefault = 2
End Enum

'----
' DXF/DWG Output formats
'----
Public Enum swDxfFormat_e

        swDxfFormat_R12 = 0
        swDxfFormat_R13 = 1
        swDxfFormat_R14 = 2
        swDxfFormat_R2000 = 3
End Enum

'---
' DXF/DWG output arrow directions
'---
Public Enum swArrowDirection_e

        swINSIDE = 0
        swOUTSIDE = 1
        swSMART = 2
End Enum


'---
' Print Properties
' The different property types for IModelDoc::SetPrintSetUp
'---
Public Enum swPrintProperties_e

        swPrintPaperSize = 0
        swPrintOrientation = 1
End Enum


'---
' Tiff Image types
'---
Public Enum swTiffImageType_e

        swTiffImageBlackAndWhite = 0
        swTiffImageRGB = 1
End Enum


'---
' Tiff Image Compression schemes
'---
Public Enum swTiffCompressionScheme_e

        swTiffUncompressed = 0
        swTiffPackbitsCompression = 1
        swTiffGroup4FaxCompression = 2
End Enum

'----
' Body operations.  For use with Body::Operations method.
'----
Public Const SWBODYINTERSECT As Integer = 15901
Public Const SWBODYCUT As Integer = 15902
Public Const SWBODYADD As Integer = 15903
Public Enum swBodyOperationError_e

        swBodyOperationUnknownError = -1
        swBodyOperationNoError = 0
        swBodyOperationNonApiBody = 1
        swBodyOperationWrongType = 2
        swBodyOperationBooleanFail = 1058
        swBodyOperationNoIntersect = 1067
        swBodyOperationNonManifold = 547
        swBodyOperationPartialCoincidence = 1040
        swBodyOperationIntersectSolidWithSheets = 972
        swBodyOperationUniteSolidSheet = 543
        swBodyOperationMissingGeom = 96
        swBodyOperationSameToolAndTarget = 545
        swBodyOperationFailGeomCondition = 3
        swBodyOperationFailToCutBody = 4
        swBodyOperationDisjointBodies = 5
        swBodyOperationEmptyBody = 6

End Enum

'---
' End Conditions.
' These are used with FeatureBoss FeatureCut, FeatureExtrusion, etc.
' Not all types are valid for all body operations.  Some of these end conditions require additional
' selections (ie - swEndCondUpToSurface etc.) and some require additional data (ie - swEndCondOffsetFromSurface)
'---
Public Enum swEndConditions_e

        swEndCondBlind = 0
        swEndCondThroughAll = 1
        swEndCondThroughNext = 2
        swEndCondUpToVertex = 3
        swEndCondUpToSurface = 4
        swEndCondOffsetFromSurface = 5
        swEndCondMidPlane = 6
End Enum

Public Enum swChamferType_e

        swChamferAngleDistance = 1
        swChamferDistanceDistance = 2
        swChamferVertex = 3
        swChamferEqualDistance = 4
End Enum

'---
' Line weights
'---
Public Enum swLineWeights_e
 
        swLW_NONE = -1
        swLW_THIN = 0
        swLW_NORMAL = 1
        swLW_THICK = 2
        swLW_THICK2 = 3
        swLW_THICK3 = 4
        swLW_THICK4 = 5
        swLW_THICK5 = 6
        swLW_THICK6 = 7
        swLW_NUMBER = 8
        swLW_LAYER = 9
End Enum

'---
' Toolbar States.  For use with ISldWorks::GetToolbarState()
'---
Public Enum swToolbarStates_e
 
        swToolbarHidden = 0
End Enum

'----
' Summary info fields for use with IModelDoc::Get/SetSummaryInfo
'----

Public Enum swSummInfoField_e

        swSumInfoTitle = 0
        swSumInfoSubject = 1
        swSumInfoAuthor = 2
        swSumInfoKeywords = 3
        swSumInfoComment = 4
        swSumInfoSavedBy = 5
        swSumInfoCreateDate = 6
        swSumInfoSaveDate = 7
        swSumInfoCreateDate2 = 8
        swSumInfoSaveDate2 = 9
End Enum

' CPropertySheet enumerated types.
' For use with the ISldWorks::PropertySheetCreateNotify notification
Public Enum swPropSheetType_e

        swPropSheetNotValid = 0
        swPropSheetLighting = 1
        swPropSheetToolsOptions = 2
        swPropSheetAmbientLight = 3
        swPropSheetDirectionalLight = 4
        swPropSheetPositionLight = 5
        swPropSheetSpotLight = 6
End Enum

Public Enum swWindowState_e

        swWindowNormal = 0
        swWindowMaximized = 1
        swWindowMinimized = 2
End Enum

' Possible values for Witness Line visibility for use by
' auDisplayDimension_c::GetWitnessVisibility and SetWitnessVisibility.
Public Enum swWitnessLineVisibility_e

        swWitnessLineBoth = 0           ' BOTH witness lines are displayed
        swWitnessLineFirst = 1          ' only FIRST witness line is displayed
        swWitnessLineSecond = 2         ' only SECOND witness line is displayed
        swWitnessLineNone = 3           ' NEITHER witness line is displayed
End Enum

' Possible values for Leader Line visibility for use by
' auDisplayDimension_c::GetLeaderVisibility and SetLeaderVisibility.
Public Enum swLeaderLineVisibility_e

        swLeaderLineBoth = 0            ' BOTH leader lines are displayed
        swLeaderLineFirst = 1           ' only FIRST leader line is displayed
        swLeaderLineSecond = 2          ' only SECOND leader line is displayed
        swLeaderLineNone = 3            ' NEITHER leader line is displayed
End Enum

' Possible values for Arrow positions for use by
' auDisplayDimension_c::GetArrowSide and SetArrowSide.
Public Enum swDimensionArrowsSide_e

        swDimArrowsInside = 0                   ' place arrows INSIDE of the witness lines
        swDimArrowsOutside = 1                  ' place arrows OUTSIDE of the witness lines
        swDimArrowsSmart = 2                    ' place arrows inside if the text and arrows fit, outside if not
        swDimArrowsFollowDoc = 3        ' place arrows the same as the document default for placing arrows
End Enum

' The different parts of the dimension text for use by
' auDisplayDimension_c::GetText and SetText.
Public Enum swDimensionTextParts_e

        swDimensionTextAll = 0                          ' all pieces of text (used only by SetText)
        swDimensionTextPrefix = 1               ' the prefix portion of the text
        swDimensionTextSuffix = 2               ' the suffix portion of the text
        swDimensionTextCalloutAbove = 3         ' the callout portion of the text, above the dimension
        swDimensionTextCalloutBelow = 4                 ' the callout portion of the text below the dimension
End Enum

Public Enum swTopology_e

        swTopoSolidBody = 1
        swTopoSheetBody = 2
        swTopoWireBody = 3
        swTopoMinimumBody = 4
End Enum

Public Enum swTopoEntity_e

        swTopoVertex = 1
        swTopoEdge = 2
        swTopoLoop = 3
        swTopoFace = 4
        swTopoShell = 5
        swTopoBody = 6
End Enum

' The alignment information possible for Views.  For use with auDrView_c::GetAlignment.
Public Enum swViewAlignment_e

        swViewAlignNone = 0                     ' this view has no alignment restrictions
        swViewAlignedChildren = 1       ' this view has children aligned with it
        swViewAligned = 2                       ' this view is aligned with a parent view
        swViewAlignBoth = 3                             ' this view is aligned and has aligned children
End Enum

'Toolbars
Public Enum swToolbar_e

                swSketchToolsToolbar = 0
                'swDependencyToolbar = 1
                swMainToolbar = 1
                swStandardToolbar = 2
                swViewToolbar = 3
                swSketchRelationsToolbar = 4
                swMacroToolbar = 5
                swSketchToolbar = 6
                swAssemblyToolbar = 7
                swDrawingToolbar = 8
                swAnnotationToolbar = 9
                swWebToolbar = 10
                swFeatureToolbar = 11
                swFontToolbar = 12
                swLineToolbar = 13

' Added for SolidWorks 99
                swSelectionFilterToolbar = 14
                swReferenceGeometryToolbar = 15
                swStandardViewsToolbar = 16
                swToolsToolbar = 17

' Added for SolidWorks 2000
                swCurvesToolbar = 18
                swMoldToolsToolbar = 19
                swSheetMetalToolbar = 20
                swSurfacesToolbar = 21
                swAlignToolbar = 22
                swLayerToolbar = 23
End Enum


' Annotations
Public Enum swInsertAnnotation_e

        swInsertCThreads = &H1
        swInsertDatums = &H2
        swInsertDatumTargets = &H4
        swInsertDimensions = &H8
        swInsertInstanceCounts = &H10
        swInsertGTols = &H20
        swInsertNotes = &H40
        swInsertSFSymbols = &H80
        swInsertWelds = &H100
        swInsertAxes = &H200
        swInsertCurves = &H400
        swInsertPlanes = &H800
        swInsertSurfaces = &H1000
        swInsertPoints = &H2000
        swInsertOrigins = &H4000
End Enum

' MessageBox values
Public Enum swMessageBoxIcon_e

        ' Icon types
        swMbWarning = 1
        swMbInformation = 2
        swMbQuestion = 3
        swMbStop = 4
End Enum

Public Enum swMessageBoxBtn_e

        ' button types
        swMbAbortRetryIgnore = 1
        swMbOk = 2
        swMbOkCancel = 3
        swMbRetryCancel = 4
        swMbYesNo = 5
        swMbYesNoCancel = 6
End Enum

Public Enum swMessageBoxResult_e

        ' return types
        swMbHitAbort = 1
        swMbHitIgnore = 2
        swMbHitNo = 3
        swMbHitOk = 4
        swMbHitRetry = 5
        swMbHitYes = 6
        swMbHitCancel = 7
End Enum

' Annotation types
Public Enum swAnnotationType_e

        swCThread = 1
        swDatumTag = 2
        swDatumTargetSym = 3
        swDisplayDimension = 4
        swGTol = 5
        swNote = 6
        swSFSymbol = 7
        swWeldSymbol = 8
        swCustomSymbol = 9
End Enum

' The possible Driven States for Dimensions.  For use with auDimension_c::DrivenState.
Public Enum swDimensionDrivenState_e

        swDimensionDrivenUnknown = 0            ' the driven/driving state is unknown
        swDimensionDriven = 1                           ' the dimension is a driven dimension
        swDimensionDriving = 2                                  ' the dimension is a driving dimension
End Enum

Public Enum swFileLoadError_e

        swGenericError = &H1
        swFileNotFoundError = &H2
        swIdMatchError = &H4
        swReadOnlyWarn = &H8
        swSharingViolationWarn = &H10
        swDrawingANSIUpdateWarn = &H20
        swSheetScaleUpdateWarn = &H40
        swNeedsRegenWarn = &H80
        swBasePartNotLoadedWarn = &H100
        swFileAlreadyOpenWarn = &H200           ' the requested file is already open, that document will be used
        swInvalidFileTypeError = &H400          ' the type argument passed into the API is not valid
        swDrawingsOnlyRapidDraftWarn = &H800    ' only drawings can be converted to RapidDraft format
        swViewOnlyRestrictions = &H1000         ' a document being opened view only can not be configured
        swFutureVersion = &H2000                ' document being opened is of a future version.
End Enum

Public Enum swFileSaveError_e

        swGenericSaveError = &H1
        swReadOnlySaveError = &H2
        swFileNameEmpty = &H4                           ' The filename must not be empty
        swFileNameContainsAtSign = &H8          ' The filename can not contain an '@' character
        swFileLockError = &H10
        swFileSaveFormatNotAvailable = &H20             ' The save as file type is not valid
        swFileSaveWithRebuildError = &H40       ' The file was saved with a rebuild error
        swFileSaveAsDoNotOverwrite = &H80       ' The user chose not to overwrite an existing file

End Enum

Public Enum swActivateDocError_e

        swGenericActivateError = &H1
        swDocNeedsRebuildWarning = &H2
End Enum

' The suppression information possible for Components.  For use with auComponent_c::Suppression.
Public Enum swComponentSuppressionState_e

        swComponentSuppressed = 0               ' Fully suppressed - nothing is loaded
        swComponentLightweight = 1              ' Featherweight - only graphics data is loaded
        swComponentFullyResolved = 2                    ' Fully resolved - model is completly loaded
End Enum

' The visibility information possible for Components.  For use with auComponent_c::Visibility.
Public Enum swComponentVisibilityState_e

        swComponentHidden = 0
        swComponentVisible = 1
End Enum

' Possible values for the solving option of components.
Public Enum swComponentSolvingOption_e

        swComponentRigidSolving = 0
        swComponentFlexibleSolving = 1
End Enum

Public Enum swCustomInfoType_e

        swCustomInfoUnknown = 0
        swCustomInfoText = 30 ' VT_LPSTR
        swCustomInfoDate = 64 ' VT_FILETIME
        swCustomInfoNumber = 3 ' VT_I4
        swCustomInfoYesOrNo = 11 ' VT_BOOL
End Enum

Public Enum swComponentResolveStatus_e

        swResolveOk = 0
        swResolveAbortedByUser = 1
        swResolveNotPerformed = 2
        swResolveError = 3
End Enum

Public Enum swSuppressionError_e

        swSuppressionBadComponent = 0
        swSuppressionBadState = 1
        swSuppressionChangeOk = 2
        swSuppressionChangeFailed = 3
End Enum

Public Enum swDynamicMode_e

        swNoDynamics = 0
        swSpinDynamics = 1
        swPanDynamics = 2
        swZoomDynamics = 3
        swUnknownDynamics = 4
        swAnimDynamics = 5
End Enum

' The justification of text with respect to the note origin.
' Used by auNote_c::Get and SetTextJustification[AtIndex].
Public Enum swTextJustification_e

        swTextJustificationLeft = 1             ' Text is Left Justified (Top Justified is assumed?)
        swTextJustificationCenter = 2           ' Text is Center Justified (Top Justified is assumed?)
        swTextJustificationRight = 3                    ' Text is Top Justified (Top Justified is assumed?)
End Enum

Public Enum swComponentReloadOption_e

        swAlwaysReload = 0
        swDontReloadOldComponents = 1
End Enum

Public Enum swComponentReloadError_e

        swReloadOkay = 0
        swWriteAccessError = 1
        swFutureVersionError = 2
        swModifiedNotReloadedError = 3
        swInvalidOption = 4
        swFileNotSavedError = 5
        swInvalidComponentError = 6
        swUnexpectedError = 7
        swComponentLightWeightError = 8
End Enum

Public Enum swIntersectionType_e

        swIntersectionSIMPLE = 1
        swIntersectionTANGENT = 2
        swIntersectionCOINCIDENCE_START = 3
        swIntersectionCOINCIDENCE_END = 4
End Enum

Public Enum swAddOrdinateDims_e

        swOrdinate = 1
        swVerticalOrdinate = 2
        swHorizontalOrdinate = 3
End Enum

Public Enum swSheetSewingOption_e

        swSewToSolid = 0
        swSewToSheets = 1
        swSewToSolidOrSheets = 2
End Enum

Public Enum swSheetSewingError_e

        swSewingOk = 0
        swBadArgument = 1
        swUnspecifiedError = 2
        swSewingFailed = 3
        swSewingIncomplete = 4
End Enum

Public Enum swBodyType_e

        swSolidBody = 0
        swSheetBody = 1
        swWireBody = 2
        swMinimumBody = 3
        swGeneralBody = 4
        swEmptyBody = 5
End Enum

' Which Configurations the Set Value applies to.
Public Enum swSetValueInConfiguration_e

        swSetValue_UseCurrentSetting = 0                ' Use the setting this parameter currently has
        swSetValue_InThisConfiguration = 1
        swSetValue_InAllConfigurations = 2
End Enum

' Return status of the Set Value operation.
Public Enum swSetValueReturnStatus_e

        swSetValue_Successful = 0
        swSetValue_Failure = 1                          ' failed for an unknown reason
        swSetValue_InvalidValue = 2             ' not a valid value for Change Parameter
        swSetValue_DrivenDimension = 3          ' can not be done on a dimension driven by geometry
        swSetValue_ModelNotLoaded = 4           ' the model must be loaded in order to set this value
End Enum

' Possible values for the bendState value for the Get/SetBendState APIs.
Public Enum swSMBendState_e

        swSMBendStateNone = 0                   ' No bend state - not a sheet metal part
        swSMBendStateSharps = 1         ' Bends are in the sharp state - bends currently not applied
        swSMBendStateFlattened = 2              ' Bends are flattened
        swSMBendStateFolded = 3                 ' Bends are fully applied
End Enum

' Possible return status of Sheet Metal APIs.
Public Enum swSMCommandStatus_e

        swSMErrorNone = 0
        swSMErrorUnknown = 1                            ' failed for an unknown reason
        swSMErrorNotAPart = 2                           ' Sheet Metal commands only apply to SW Parts
        swSMErrorNotASheetMetalPart = 3 ' the part contains no Sheet Metal features
        swSMErrorInvalidBendState = 4           ' an invalid bend state was specified (Set Bend State)
End Enum

' Feature error code returned from Feature::GetErrorCode.
Public Enum swFeatureError_e

        swFeatureErrorNone = 0                                                  ' No error
        swFeatureErrorUnknown = 1                                               ' Unknown error
                
        swFeatureErrorFilletNoLoop = 10                         ' Loop for fillet/chamfer does not exist
        swFeatureErrorFilletNoFace = 11                         ' face for fillet/chamfer does not exist
        swFeatureErrorFilletInvalidRadius = 12                  ' invalid fillet radius or a face blend fillet recommended
        swFeatureErrorFilletNoEdge = 13                         ' Edge for fillet/chamfer does not exist
        swFeatureErrorFilletModelGeometry = 14                  ' Failed to create fillet due to model geometry
        swFeatureErrorFilletRadiusTooSmall = 15         ' Radius value is too small
        swFeatureErrorFilletCannotExtend = 16                   ' Selected elements cannot be extended to intersect
        swFeatureErrorFilletRadiusEliminateElement = 17 ' Specified radius would eliminate one of the elements
        swFeatureErrorFilletRadiusTooBig = 18                   ' Radius is too big or the elements are tangent or nearly tangent
        swFeatureErrorFilletRadiusTooBig2 = 19                  ' The radius of the fillet is too large to fit the surrounding geometry. Try adjusting the input geometry and radius values or try using a face blend fillet.
                
        swFeatureErrorExtrusionDisjoint = 30                                    ' This feature would create a disjoint body. The direction may be wrong
        swFeatureErrorExtrusionNoEndFound = 31                                  ' Cannot locate end of feature
        swFeatureErrorExtrusionBadGeometricConditions = 32              ' Unable to create this extruded feature due to geometric conditions
        swFeatureErrorExtrusionCutContourOpenAndClosed = 33     ' Extruded cuts cannot have both open and closed contours
        swFeatureErrorExtrusionCutContourInvalid = 34                   ' Extruded cuts require at least one closed or open contour which does not self-intersect
        swFeatureErrorExtrusionOpenCutContourInvalid = 35               ' Open extruded cuts require a single open contour which does not self-intersect
        swFeatureErrorExtrusionBossContourOpenAndClosed = 36    ' Bosses cannot have both open and closed contours
        swFeatureErrorExtrusionBossContourInvalid = 37                  ' Bosses require one or more closed contours which do not self-intersect
                
        'To be continued...
End Enum

' Possible values for the saveAsVersion argument of the SaveAs2 API
Public Enum swSaveAsVersion_e

        swSaveAsCurrentVersion = 0              ' default

' Support for Save As 98plus was added and removed during the Sw99 development cycle.  It is no longer
' a possibility however existing programs / macros could be trying to use it. (SPR 61060)
        swSaveAsSW98plus = 1                    ' save SW model in SW98plus model format - NO LONGER SUPPORTED

        swSaveAsFormatProE = 2                  ' save Sw part as Pro/E format .prt/.asm extension (not as Sw .prt/.asm)
End Enum

' Possible values for the leaderType argument of the Get/SetArcLengthLeader APIs
Public Enum swArcLengthLeaderType_e

        swArcLengthLeaderParallel = 1   ' Leaders are parallel to each other
        swArcLengthLeaderRadial = 2             ' Leaders are radial from the arc center point
End Enum

' Possible values for the condition argument of the Get/SetArcEndCondition APIs
' These values should match up with the values in the moArcDimType_e enumeration.
Public Enum swArcEndCondition_e

        swArcEndConditionNone = 0               ' The end point is not related to an arc
        swArcEndConditionCenter = 1     ' The end point is the center of the arc
        swArcEndConditionMin = 2                ' The end point is the nearest point on the arc
        swArcEndConditionMax = 3                ' The end point is the furthest point on the arc
End Enum

Public Enum swDestroyNotifyType_e

        swDestroyNotifyDestroy = 0              ' The view is being destroyed
        swDestroyNotifyHidden = 1               ' The view is actually being hidden not destroyed
End Enum

Public Enum swSketchSegments_e

        swSketchLINE = 0
        swSketchARC = 1
        swSketchELLIPSE = 2
        swSketchSPLINE = 3
        swSketchTEXT = 4
        swSketchPARABOLA = 5
End Enum

Public Enum swPipingPenetrationStatus_e

        swPenetrationSucceeded = 0
        swPenetrationFailed = 1
        swPenetrationFailedPipeTooWide = 2
        swPenetrationFailedDllNotLoaded = 3
        swPenetrationFailedNoSelection = 4
        swPenetrationFailedNotRouting = 5
        swPenetrationFailedBadSelection = 6
        swPenetrationFailedBadFitting = 7
        swPenetrationFailedAlreadyPenetrating = 8
End Enum

' Enumerate the possible different entity types for passing as a notification argument.
' Currently used by AddItemNotify RenameItemNotify, and DeleteItemNotify.
Public Enum swNotifyEntityType_e

        swNotifyConfiguration = 1                       ' Configuration is being added, renamed, or deleted
        swNotifyComponent = 2
End Enum

Public Enum swRayPtsOpts_e

        swRayPtsOptsNORMALS = &H1
        swRayPtsOptsTOPOLS = &H2
        swRayPtsOptsENTRY_EXIT = &H4
        swRayPtsOptsUNBLOCK = &H8       'alow the system to respond while waiting
End Enum

Public Enum swRayPtsResults_e

        swRayPtsResultsFACE = &H1
        swRayPtsResultsSILHOUETTE = &H2
        swRayPtsResultsEDGE = &H4
        swRayPtsResultsVERTEX = &H8
        swRayPtsResultsENTER = &H10
        swRayPtsResultsEXIT = &H20
End Enum

' The different pieces of text within a weld annotation.  (WeldSymbol::GetText)
Public Enum swWeldSymbolTextTypes_e

        swWeldLeftTextAbove = 1         ' The text just to the left of the weld symbol, above the horizontal line
        swWeldSymbolTextAbove = 2               ' The weld symbol, above the horizontal line
        swWeldRightTextAbove = 3                ' The text just to the right of the weld symbol, above the horizontal line
        swWeldStaggerTextAbove = 4              ' The text related to the stagger characteristic, above the horizontal line
        swWeldLeftTextBelow = 5         ' The text just to the left of the weld symbol, below the horizontal line
        swWeldSymbolTextBelow = 6               ' The weld symbol, below the horizontal line
        swWeldRightTextBelow = 7                ' The text just to the right of the weld symbol, below the horizontal line
        swWeldStaggerTextBelow = 8              ' The text related to the stagger characteristic, below the horizontal line
        swWeldProcessText = 9                   ' The text related to the process indicators characteristic
End Enum

' The different cases for contour symbols of a weld annotation.  (WeldSymbol::GetContour/SetText)
Public Enum swWeldSymbolContourTypes_e

        swWeldContourNone = 1
        swWeldContourFlat = 2
        swWeldContourConvex = 3
        swWeldContourConcave = 4
End Enum

' The different cases for symmetric characteristic of a weld annotation. (WeldSymbol::Get/SetSymmetric)
Public Enum swWeldSymbolSymmetric_e

        swWeldSymmetric = 1                     ' The symbol is symmetric on this weld annotation
        swWeldDashedLineOnTop = 2               ' The symbol is not symmetric, with the dashed horizontal line above
        swWeldDashedLineOnBottom = 3    ' The symbol is not symmetric with the dashed horizontal line below
End Enum

' The different cases for field or site characteristic of a weld annotation. (WeldSymbol::Get/SetFieldWeld)
Public Enum swWeldSymbolField_e

        swFieldWeldNone = 1                     ' No field-site weld marking on this annotation
        swFieldWeldUp = 2                               ' The field-site weld marking is pointing up
        swFieldWeldDown = 3                             ' The field-site weld marking is pointing down
End Enum

' The different cases for whether or not a Display Dimension leader is broken or not and how the text is
' placed relative to the leader.
Public Enum swDisplayDimensionLeaderText_e

        swSolidLeaderAlignedText = 1
        swBrokenLeaderHorizontalText = 2
        swBrokenLeaderAlignedText = 3
End Enum

Public Enum swLineStyles_e

        swLineCONTINUOUS = 0
        swLineHIDDEN = 1
        swLinePHANTOM = 2
        swLineCHAIN = 3
        swLineCENTER = 4
        swLineSTITCH = 5
        swLineCHAINTHICK = 6
End Enum

' The different types of drawing views.
Public Enum swDrawingViewTypes_e

        swDrawingSheet = 1
        swDrawingSectionView = 2
        swDrawingDetailView = 3
        swDrawingProjectedView = 4
        swDrawingAuxiliaryView = 5
        swDrawingStandardView = 6
        swDrawingNamedView = 7
        swDrawingRelativeView = 8
End Enum

' For the Sketch Fillet command the different actions that can be taken when the fillet is being
' applied to a corner that has constraints.
Public Enum swConstrainedCornerAction_e

        swConstrainedCornerInteract = 0         ' Ask the user whether to Delete the geometry or Stop Processing
        swConstrainedCornerKeepGeometry = 1     ' Keep the constrained geometry in the part
        swConstrainedCornerDeleteGeometry = 2   ' Delete the constrained geometry from the part
        swConstrainedCornerStopProcessing = 3   ' Do not do anything stop processing immediately
End Enum

' The different command mode that can be in effect.  Used by the SldWorks::GetMouseDragMode API.
Public Enum swMouseDragMode_e

        swTranslateAssemblyComponent = 1                                ' Assembly Component Move mode
        swRotateAssemblyComponentAboutCenter = 2                ' Assembly Component Rotate mode
        swRotateAssemblyComponentAboutAxis = 3                  ' Assembly Component Rotate About Axis mode
        swAssemblySmartMates = 4                                                ' Assembly Component Smart Mate mode
        swRotateView = 5                                                                ' View Rotate mode
        swTranslateView = 6                                                     ' View Translate mode
        swZoomView = 7                                                                  ' View Zoom mode
        swZoomToAreaOfView = 8                                                  ' View Zoom To mode
        swInsertDimension = 9                                                   ' Insert Dimension mode
End Enum

' The different possibilities for types of Datum Target area shapes. (DatumTargetSym::Get/SetTargetShape)
Public Enum swDatumTargetAreaShape_e

        swDatumTargetAreaNone = 0
        swDatumTargetAreaPoint = 1
        swDatumTargetAreaCircle = 2
        swDatumTargetAreaRectangle = 3
End Enum

' Possible status values for the Edit Part command.
Public Enum swEditPartCommandStatus_e

' The values < 0 indicate a complete failure of the Edit Part command and the reason for failure.
        swEditPartFailure = -1
        swEditPartAsmMustBeSaved = -2
        swEditPartCompMustBeSelected = -3
        swEditPartCompMustBeResolved = -4
        swEditPartCompMustHaveWriteAccess = -5

' The values >= 0 indicate that the Edit Part command was successful.  The values > 0 indicate any additional
' information that might be important to the caller the kind of information that might be presented to the user
' in a message box if this is an interactive user.  This information is probably most important to an API user.
        swEditPartSuccessful = 0
        swEditPartCompNotPositioned = &H1
End Enum

' Possible values for the visibility state of an annotation. (Annotation::Visible property)
Public Enum swAnnotationVisibilityState_e

        swAnnotationVisibilityUnknown = 0
        swAnnotationVisible = 1
        swAnnotationHalfHidden = 2
        swAnnotationHidden = 3
End Enum

' This is used by the notification LightweightComponentOpenNotify()
Public Enum swOutOfDateStatus_e

        swUnknownState = 0
        swModelUpToDate = 1
        swModelOutOfDate = 2
End Enum

' This is used by the API GetLocalizedMenuName()
Public Enum swMenuIdentifiers_e

        swFileMenu = 0
        swEditMenu = 1
        swViewMenu = 2
        swInsertMenu = 3
        swToolsMenu = 4
        swWindowMenu = 5
        swHelpMenu = 6
        swDeveloperToolsMenu = 7
        swViewToolbarsMenu = 8
End Enum

' For InsertScale
Public Enum swScaleType_e

        swScaleAboutCentroid = 0
        swScaleAboutOrigin = 1
        swScaleAboutCoordinateSystem = 2
End Enum

' For InsertCavity4
Public Enum swCavityScaleType_e

        swAboutCentroid = 0
        swAboutOrigin = 1
        swAboutMoldBaseOrigin = 2
        swAboutCoordinateSystem = 3
End Enum

Public Enum swFeatMgrPane_e

        swFeatMgrPaneTop = 0
        swFeatMgrPaneBottom = 1
        swFeatMgrPaneTopHidden = 2
        swFeatMgrPaneBottomHidden = 3
End Enum

' Possible values for the swDetailingDualDimPosition User Preference setting.
Public Enum swDetailingDualDimPosition_e

        swDualDimensionsSideBySide = 1
        swDualDimensionsAboveAndBelow = 2
End Enum

' Possible values for the swDetailingDimTrailingZero User Preference setting
Public Enum swDetailingDimTrailingZero_e

        swDimSmartTrailingZeroes = 0
        swDimShowTrailingZeroes = 1
        swDimRemoveTrailingZeroes = 2
End Enum

' Possible values for the swDetailingToleranceTextSizing User Preference setting
Public Enum swDetailingToleranceTextSizing_e

        swToleranceTextSizeUsingScaleValue = 1
        swToleranceTextSizeUsingHeightValue = 2
End Enum

' Possible values for the swDetailingDimensionStandard User Preference setting
Public Enum swDetailingStandard_e

        swDetailingStandardANSI = 1
        swDetailingStandardISO = 2
        swDetailingStandardDIN = 3
        swDetailingStandardJIS = 4
        swDetailingStandardBS = 5
        swDetailingStandardGOST = 6
End Enum

' Possible values for the swDetailingBOMUpperText and LowerText User Preference settings
Public Enum swDetailingNoteTextContent_e

        swDetailingNoteTextCustom = 1
        swDetailingNoteTextItemNumber = 2
        swDetailingNoteTextQuantity = 3
End Enum

' Possible values for the swDetailingVirtualSharpStyle User Preference settings
Public Enum swDetailingVirtualSharp_e

        swDetailingVirtualSharpNone = 0
        swDetailingVirtualSharpPlus = 1
        swDetailingVirtualSharpStar = 2
        swDetailingVirtualSharpWitness = 3
        swDetailingVirtualSharpDot = 4
End Enum

' Different types of dimensions.  Used by DisplayDimension::GetType.
Public Enum swDimensionType_e

        swDimensionTypeUnknown = 0
        swOrdinateDimension = 1
        swLinearDimension = 2
        swAngularDimension = 3
        swArcLengthDimension = 4
        swRadialDimension = 5
End Enum

' Possible values for the swImageQualityShaded User Preference setting
Public Enum swImageQualityShaded_e

        swShadedImageQualityCoarse = 1
        swShadedImageQualityFine = 2
        swShadedImageQualityCustom = 3
End Enum

' Possible values for the swImageQualityWireframe User Preference setting
Public Enum swImageQualityWireframe_e

        swWireframeImageQualityOptimal = 1
        swWireframeImageQualityCustom = 2
End Enum

Public Enum swLoadDetachedModelRules_e

        swLoadDetachedModelPrompt = 0
        swLoadDetachedModelAuto = 1
        swDoNotLoadDetachedModel = 2
End Enum

' Possible value for different methods of display tangent edges.  (View::Get/SetDisplayTangentEdges2)
Public Enum swDisplayTangentEdges_e

        swTangentEdgesHidden = 0
        swTangentEdgesVisibleAndFonted = 1
        swTangentEdgesVisible = 2
End Enum

' Possible values for the swSTLQuality User Preference setting
Public Enum swSTLQuality_e

        swSTLQuality_Coarse = 1
        swSTLQuality_Fine = 2
        swSTLQuality_Custom = 3
End Enum

' Possible values for the swDrawingProjectionType User Preference setting
Public Enum swDrawingProjectionType_e

        swDrawing1stAngleProjection = 1
        swDrawing3rdAngleProjection = 2
End Enum

Public Enum swPromptAlwaysNever_e

        swResponsePrompt = 0
        swResponseAlways = 1
        swResponseNever = 2
End Enum

' Possible values for the swIGESRepresentation User Preference setting.
Public Enum swIGESRepresentation_e

        swIGES_TRMSRF = 0               ' Trimmed surface representation
        swIGES_CURVES = 1               ' WireFrame representation
End Enum

' Possible values for the swIGESSystem User Preference setting.
Public Enum swIGESPreferredSystem_e

        swIGES_STANDARD = 0
        swIGES_NURBS = 1
        swIGES_ANSYS = 2
        swIGES_COSMOS = 3
        swIGES_MASCAM = 4
        swIGES_SURFCAM = 5
        swIGES_SMARTCAM = 6
        swIGES_TEKSOFT = 7
        swIGES_ALPHACAM = 8
        swIGES_MULTICAM = 9
End Enum

' Possible values for the swIGESCurveRepresentation User Preference setting.
Public Enum swIGESCurveRepresentation_e

        swIGES_CURVES_BSPLINE = 0               ' free form curves as bspline representation
        swIGES_CURVES_PSPLINE = 1               ' free form curves as parametric spline representation
End Enum

'Possible value for the constraint status of a sketch
Public Enum swConstrainedStatus_e

        swUnknownConstraint = 1
        swUnderConstrained = 2
        swFullyConstrained = 3
        swOverConstrained = 4
        swNoSolution = 5
        swInvalidSolution = 6
        swAutosolveOff = 7
End Enum

' Suppression actions for features.
Public Enum swFeatureSuppressionAction_e

        swSuppressFeature = 0                   ' Suppress the feature.
        swUnSuppressFeature = 1                 ' Unsuppress the feature.
        swUnSuppressDependent = 2               ' Unsuppress the children of the features.
End Enum

' HLR Quality Settings...
Public Enum swHlrQuality_e

        swPreciseHlr = 0
        swFastHlr = 1
End Enum

' Possible values for the entityType argument of the Sketch::SetEntityCount API.
Public Enum swSketchEntityType_e

        swSketchEntityPoint = 1
        swSketchEntityLine = 2
        swSketchEntityArc = 3
        swSketchEntityEllipse = 4
        swSketchEntityParabola = 5
        swSketchEntitySpline = 6
End Enum

Public Enum swWzdHoleTypes_e

        swSimple = 0
        swTapered = 1
        swCounterBored = 2
        swCounterSunk = 3
        swCounterDrilled = 4
        swSimpleDrilled = 5
        swTaperedDrilled = 6
        swCounterBoredDrilled = 7
        swCounterSunkDrilled = 8
        swCounterDrilledDrilled = 9

        swCounterBoreBlind = 10
        swCounterBoreBlindCounterSinkMiddle = 11
        swCounterBoreBlindCounterSinkTop = 12
        swCounterBoreBlindCounterSinkTopmiddle = 13
        swCounterBoreThru = 14
        swCounterBoreThruCounterSinkBottom = 15
        swCounterBoreThruCounterSinkMiddle = 16
        swCounterBoreThruCounterSinkMiddleBottom = 17
        swCounterBoreThruCounterSinkTop = 18
        swCounterBoreThruCounterSinkTopBottom = 19
        swCounterBoreThruCounterSinkTopMiddle = 20
        swCounterBoreThruCounterSinkTopMiddleBottom = 21

        swHoleBlind = 22
        swHoleBlindCounterSinkTop = 23

        swCounterSinkBlind = 24

        swHoleThru = 25
        swHoleThruCounterSinkBottom = 26
        swHoleThruCounterSinkTop = 27
        swHoleThruCounterSinkTopBottom = 28

        swCounterSinkThru = 29
        swCounterSinkThruCounterSinkBottom = 30

        swTapBlind = 31
        swTapBlindCounterSinkTop = 32
        swTapThru = 33
        swTapThruCounterSinkBottom = 34
        swTapThruCounterSinkTop = 35
        swTapThruCounterSinkTopBottom = 36

        swPipeTapBlind = 37
        swPipeTapBlindCounterSinkTop = 38
        swPipeTapThru = 39
        swPipeTapThruCounterSinkBottom = 40
        swPipeTapThruCounterSinkTop = 41
        swPipeTapThruCounterSinkTopBottom = 42
End Enum

' Update this when you add new hole types.
Public Const NUM_HOLE_TYPES As Integer = 43

Public Enum swCreateFacesBodyAction_e
    swCreateFacesBodyActionCap = 1
    swCreateFacesBodyActionGrow = 2
    swCreateFacesBodyActionGrowFromParent = 3
    swCreateFacesBodyActionLeaveRubber = 4
End Enum

' Used for APIs that allow bitwise ORing of document types like
' SldWorks::AddToolbar2()
Public Enum swDocTemplateTypes_e
    swDocTemplateTypeNONE = &H1
    swDocTemplateTypePART = &H2
    swDocTemplateTypeASSEMBLY = &H4
    swDocTemplateTypeDRAWING = &H8
End Enum

Public Enum swCreateFeatureBodyOpts_e

        swCreateFeatureBodyCheck = &H1
        swCreateFeatureBodySimplify = &H2
End Enum

Public Enum swToolbarDockStatePosition_e

    swDockNoToolbar = -1
    swNoDock = 0
    swDockTop = 1
    swDockBottom = 2
    swDockRight = 3
    swDockLeft = 4
End Enum

Public Enum swImprintingFacesOpts_e

    swImprintingFacesOnTool = &H1
    swImprintingFacesOnOverlapping = &H2
    swImprintingFacesOnExtendFace = &H4
End Enum

' A list of feature types to be used with the Sketch::CheckFeatureUse API
Public Enum swSketchCheckFeatureProfileUsage_e

        swSketchCheckFeature_UNSET = 0
        swSketchCheckFeature_BASEEXTRUDE = 1
        swSketchCheckFeature_BASEEXTRUDETHIN = 2
        swSketchCheckFeature_BOSSEXTRUDE = 3
        swSketchCheckFeature_BOSSEXTRUDETHIN = 4
        swSketchCheckFeature_SURFACEEXTRUDE = 5
        swSketchCheckFeature_BASEREVOLVE = 6
        swSketchCheckFeature_BASEREVOLVETHIN = 7
        swSketchCheckFeature_BOSSREVOLVE = 8
        swSketchCheckFeature_BOSSREVOLVETHIN = 9
        swSketchCheckFeature_SURFACEREVOLVE = 10
        swSketchCheckFeature_CUTEXTRUDE = 11
        swSketchCheckFeature_CUTEXTRUDETHIN = 12
        swSketchCheckFeature_CUTREVOLVE = 13
        swSketchCheckFeature_CUTREVOLVETHIN = 14
        swSketchCheckFeature_SWEEPSECTION = 15
        swSketchCheckFeature_SURFACESWEEPSECTION = 16
        swSketchCheckFeature_SWEEPPATHORGUIDE = 17
        swSketchCheckFeature_LOFTSECTION = 18
        swSketchCheckFeature_SURFACELOFTSECTION = 19
        swSketchCheckFeature_LOFTGUIDE = 20
        swSketchCheckFeature_RIBSECTION = 21
        swSketchCheckFeature_SHEETMETAL_BASEFLANGE = 22
End Enum

' A list of return status values for the Sketch::CheckFeatureUse API
Public Enum swSketchCheckFeatureStatus_e
 
        swSketchCheckFeatureStatus_UnknownError = -1
        swSketchCheckFeatureStatus_OK = 0
        swSketchCheckFeatureStatus_EntXEnt = 1
        swSketchCheckFeatureStatus_EntXSelf = 2
        swSketchCheckFeatureStatus_EntUnspecBad = 3
        swSketchCheckFeatureStatus_ThreeEnts = 4
        swSketchCheckFeatureStatus_EmptySketch = 5
        swSketchCheckFeatureStatus_WrongOpen = 6
        swSketchCheckFeatureStatus_WrongManyContours = 7
        swSketchCheckFeatureStatus_ZeroLengthEnt = 8
        swSketchCheckFeatureStatus_ManyOpen = 9
        swSketchCheckFeatureStatus_NoOpen = 10
        swSketchCheckFeatureStatus_MixedContours = 11
        swSketchCheckFeatureStatus_CturXCtur = 12
        swSketchCheckFeatureStatus_DisjCturs = 13
        swSketchCheckFeatureStatus_OpenWantClosed = 14
        swSketchCheckFeatureStatus_ClosedWantOpen = 15
        swSketchCheckFeatureStatus_DoubleContainment = 16
        swSketchCheckFeatureStatus_MoreThanOneContour = 17
        swSketchCheckFeatureStatus_OneOpenContourExpected = 18
        swSketchCheckFeatureStatus_OneClosedContourExpected = 19
        swSketchCheckFeatureStatus_WantSingleOpenOrMultiClosedDisjoint = 20
        swSketchCheckFeatureStatus_NeedsAxis = 21
        swSketchCheckFeatureStatus_OpenOrUnclear = 22
        swSketchCheckFeatureStatus_ContourIntersectsCenterLine = 23
End Enum

' A list of return status values for the ModelDoc::GetMassProperties API
Public Enum swMassPropertiesStatus_e

        swMassPropertiesStatus_OK = 0
        swMassPropertiesStatus_UnknownError = 1
        swMassPropertiesStatus_NoBody = 2
End Enum

' A list of possible arc types when using CreateTangentArc2

Public Enum swTangentArcTypes_e

        swForward = 1
        swLeft = 2
        swBack = 3
        swRight = 4
End Enum

' Possible values for the options argument of SldWorks::OpenDoc4.
Public Enum swOpenDocOptions_e

        swOpenDocOptions_Silent = &H1                   ' Open document silently or not
        swOpenDocOptions_ReadOnly = &H2         ' Open document read only or not
        swOpenDocOptions_ViewOnly = &H4         ' Open document view only or not
        swOpenDocOptions_RapidDraft = &H8               ' Convert document to RapidDraft format or not (drawings only)
        swOpenDocOptions_LoadModel = &H10               ' Load detached models automatically or not (drawings only)
End Enum

' Possible values for the options argument of ModelDoc::SaveAs3.
Public Enum swSaveAsOptions_e

    swSaveAsOptions_Silent = &H1           ' Save document silently or not
    swSaveAsOptions_Copy = &H2             ' Save document as a copy or not
    swSaveAsOptions_SaveReferenced = &H4    ' Save referenced documents or not (drawings and parts only)
End Enum

Public Enum swInConfigurationOpts_e

    swThisConfiguration = 1
    swAllConfiguration = 2
    swSpecifyConfiguration = 3
End Enum

Public Enum swKernelErrorCode_e

        swErrorSuccess = 1
        swErrorError = 0
        swErrorNotEntity = -100022
        swErrorInvalidParameter = -100120
        swErrorSurfaceDiscontinuous = -100129
        swErrorCurveDiscontinuous = -100131
        swErrorInvalidEntity = -100914
        swErrorInvalidSharing = -100921
        swErrorInvalidKnots = -100978
        swErrorInvalidGeometry = -100999
        swErrorHasInvalidentity = -101004
        swErrorBodyDontKnit = -101041
        swErrorInvalidPattern = -101042
        swErrorCurveShort = -101057
        swErrorFailed = -101063
        swErrorCheckFailed = -105061
        swErrorGeometryMissing = -113803
        swErrorTopologySelfx = -113804
        swErrorGeometrySelfx = -113805
        swErrorGeometryDegenerate = -113806
        swErrorInvalidGeometry2 = -113808
        swErrorCheckFailed2 = -113812
        swErrorFaceFaceInconsistent = -113816
        swErrorVertexNotOnCurve = -113818
        swErrorVerticesTouch = -113821
        swErrorLoopsInconsistent = -113826
        swErrorGeometryDiscontinuous = -113827
        swErrorFacecheckFailed = -113829
        swErrorFaceRedundant = -116402
        swErrorInconsistentDirs = -116403
        swErrorEdgeisectInvalid = -116404
        swErrorInvalidLoop = -116405
        swErrorEdgeIncorrectOrder = -116406
        swErrorUnknown = -1
End Enum

' Different buttons that can be displayed on the PropertyManagerPage.
Public Enum swPropertyManagerButtonTypes_e

        swPropertyManager_OkayButton = &H1
        swPropertyManager_CancelButton = &H2
End Enum

' Return status values for the various PropertyManagerPage APIs.
Public Enum swPropertyManagerStatus_e

        swPropertyManagerStatus_Okay = 0
        swPropertyManagerStatus_Failed = -1
        swPropertyManagerStatus_Disconnected = -2
End Enum

' Possible values for the swParasolidOutputVersion User Preference setting.
Public Enum swParasolidOutputVersion_e

        swParasolidOutputVersion_latest = 0
        swParasolidOutputVersion_80 = 1
        swParasolidOutputVersion_90 = 2
        swParasolidOutputVersion_91 = 3
        swParasolidOutputVersion_100 = 4
        swParasolidOutputVersion_110 = 5
        swParasolidOutputVersion_111 = 6
        swParasolidOutputVersion_120 = 7
End Enum

' Possible values for what action to take when setting the selected object mark.
Public Enum swSelectionMarkAction_e

        swSelectionMarkSet = 0
        swSelectionMarkAppend = 1
        swSelectionMarkRemove = 2
        swSelectionMarkClear = 3
End Enum

' Possible values for the swEdgesHiddenEdgeDisplay integer user preference option
Public Enum swEdgesHiddenEdgeDisplay_e

        swEdgesHiddenEdgeDisplaySolid = 1
        swEdgesHiddenEdgeDisplayDashed = 2
End Enum

' Possible values for the swEdgesTangentEdgeDisplay integer user preference option
Public Enum swEdgesTangentEdgeDisplay_e

        swEdgesTangentEdgeDisplayVisible = 1
        swEdgesTangentEdgeDisplayPhantom = 2
        swEdgesTangentEdgeDisplayRemoved = 3
End Enum

' Possible values for the swEdgesShadedModeDisplay integer user preference option
Public Enum swEdgesShadedModeDisplay_e

        swEdgesShadedModeDisplayNone = 1
        swEdgesShadedModeDisplayHLR = 2
        swEdgesShadedModeDisplayWireframe = 3
End Enum

Public Enum swSplitFaceOnParam_e

        swSplitFaceOnParamU = 1
        swSplitFaceOnParamV = 2
End Enum

Public Enum swTbCommand_e
        swTbCONTROL = -2
        swTbACTIVE = -1
        swTbNONE = 0
        swTbPART = 1
        swTbASSEMBLY = 2
        swTbDRAWING = 3
End Enum

Public Enum swTbSaveModes_e
        swTbSAVE = 0
        swTbLOAD = 1
End Enum

Public Enum swTbControlModes_e
        swTbSTOP = 0
        swTbCONTINUE = 1
        swTbOleInplaceMode = 2
End Enum

Public Enum swBendAllowanceTypes_e
        swBendAllowanceBendTable = 1
        swBendAllowanceKFactor = 2
        swBendAllowanceDirect = 3
End Enum

Public Enum swSheetMetalReliefTypes_e
        swSheetMetalReliefRectangular = 1
        swSheetMetalReliefTear = 2
        swSheetMetalReliefObround = 3
        swSheetMetalReliefNone = 4
End Enum

Public Enum swUserUnitsType_e

        swLengthUnit = 0
        swAngleUnit = 1
End Enum

Public Enum swFlangeOffsetTypes_e
        swFlangeOffsetBlind = 1
        swFlangeOffsetUpToVertex = 2
        swFlangeOffsetUpToSurface = 3
        swFlangeOffsetFromSurface = 4
        swFlangeOffsetMidPlane = 5
End Enum

Public Enum swFlangeDimTypes_e
        swFlangeDimTypeOuterVirtualSharp = 1
        swFlangeDimTypeInnerVirtualSharp = 2
End Enum

Public Enum swFlangePositionTypes_e
        swFlangePositionTypeMaterialInside = 1
        swFlangePositionTypeMaterialOutside = 2
        swFlangePositionTypeBendOutside = 3
        swFlangePositionTypeBendCenterLine = 4
End Enum

Public Enum swReliefTearTypes_e
        swReliefTearTypeRip = 1
        swReliefTearTypeExtend = 2
End Enum

Public Enum swClosedCornerTypes_e
        swClosedCornerTypeButt = 1
        swClosedCornerTypeOverlap = 2
        swClosedCornerTypeUnderlap = 3
End Enum


Public Enum swSelectionReferenceTypes_e
        swReferenceTypeVertex = 1
        swReferenceTypeEdge = 2
        swReferenceTypeFace = 3
        swReferenceTypeRefSurface = 4
        swReferenceTypeRefPlan = 5
End Enum

Public Enum swPatternReferenceTypes_e
        swPatternReferenceTypeAxis = 0
        swPatternReferenceTypeEdge = 1
        swPatternReferenceTypeRefDim = 2
End Enum


'
' To combine options for a specific API call you can use bitwise addition
' on the numbers within a particular enum.
'

' The following Public Enum represents the option bits that can be set
' for FeatureRevolve2 and FeatureCutRevolve2

Public Enum swRevolveOptions_e
        swAutoCloseSketch = &H1
End Enum

' The following Public Enum represents the option bits that can be set
' for AddConfiguration and EditConfiguration

Public Enum swConfigurationOptions_e
        swUseAlternateName = &H1
        swDontShowPartsInBOM = &H2
End Enum

' The following Public Enum represents the option bits that can be set
' for SetBlockingState

Public Enum swBlockingStates_e
        swNoBlock = &H0
        swFullBlock = &H1
        swModifyBlock = &H2
        swPartialModifyBlock = &H3
End Enum


' The following Public Enum can be used with the ModelDoc::Rebuild function.  Be aware that
' certain options are only valid for particular document types.  For example
' swUpdateMates is only valid for ModelDoc objects which are assemblies.

Public Enum swRebuildOptions_e
        swRebuildAll = &H1
        swForceRebuildAll = &H2
        swUpdateMates = &H4
        swCurrentSheetDisp = &H8
        swUpdateDirtyOnly = &H10
End Enum

Public Enum swThinWallType_e
    swThinWallOneDirection = 0
    swThinWallOppDirection = 1
    swThinWallMidPlane = 2
    swThinWallTwoDirection = 3
End Enum

' Possible values for the swPageSetupPrinterOrientation integer user preference value.
Public Enum swPageSetupOrientation_e
    swPageSetupOrient_Portrait = 1
    swPageSetupOrient_Landscape = 2
End Enum

' Possible values for the swPageSetupPrinterDrawingColor integer user preference value.
Public Enum swPageSetupDrawingColor_e
    swPageSetup_AutomaticDrawingColor = 1
    swPageSetup_ColorGrey = 2
    swPageSetup_BlackAndWhite = 3
End Enum

