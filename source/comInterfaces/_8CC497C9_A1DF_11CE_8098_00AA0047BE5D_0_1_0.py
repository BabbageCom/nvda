# -*- coding: mbcs -*-
typelib_path = 'msftedit.dll'
_lcid = 0 # change this if required
from ctypes import *
from ctypes.wintypes import _ULARGE_INTEGER
import comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0
from comtypes import GUID
from ctypes import HRESULT
from comtypes import helpstring
from comtypes import COMMETHOD
from comtypes import dispid
from ctypes.wintypes import _LARGE_INTEGER
from comtypes import BSTR
from comtypes.automation import VARIANT
from comtypes import IUnknown
from ctypes.wintypes import _FILETIME
WSTRING = c_wchar_p
from comtypes.automation import VARIANT


class ISequentialStream(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{0C733A30-2A1C-11CE-ADE5-00AA0044773D}')
    _idlflags_ = []
class IStream(ISequentialStream):
    _case_insensitive_ = True
    _iid_ = GUID('{0000000C-0000-0000-C000-000000000046}')
    _idlflags_ = []
ISequentialStream._methods_ = [
    COMMETHOD([], HRESULT, 'RemoteRead',
              ( ['out'], POINTER(c_ubyte), 'pv' ),
              ( ['in'], c_ulong, 'cb' ),
              ( ['out'], POINTER(c_ulong), 'pcbRead' )),
    COMMETHOD([], HRESULT, 'RemoteWrite',
              ( ['in'], POINTER(c_ubyte), 'pv' ),
              ( ['in'], c_ulong, 'cb' ),
              ( ['out'], POINTER(c_ulong), 'pcbWritten' )),
]
################################################################
## code template for ISequentialStream implementation
##class ISequentialStream_Impl(object):
##    def RemoteRead(self, cb):
##        '-no docstring-'
##        #return pv, pcbRead
##
##    def RemoteWrite(self, pv, cb):
##        '-no docstring-'
##        #return pcbWritten
##

class tagSTATSTG(Structure):
    pass
IStream._methods_ = [
    COMMETHOD([], HRESULT, 'RemoteSeek',
              ( ['in'], _LARGE_INTEGER, 'dlibMove' ),
              ( ['in'], c_ulong, 'dwOrigin' ),
              ( ['out'], POINTER(_ULARGE_INTEGER), 'plibNewPosition' )),
    COMMETHOD([], HRESULT, 'SetSize',
              ( ['in'], _ULARGE_INTEGER, 'libNewSize' )),
    COMMETHOD([], HRESULT, 'RemoteCopyTo',
              ( ['in'], POINTER(IStream), 'pstm' ),
              ( ['in'], _ULARGE_INTEGER, 'cb' ),
              ( ['out'], POINTER(_ULARGE_INTEGER), 'pcbRead' ),
              ( ['out'], POINTER(_ULARGE_INTEGER), 'pcbWritten' )),
    COMMETHOD([], HRESULT, 'Commit',
              ( ['in'], c_ulong, 'grfCommitFlags' )),
    COMMETHOD([], HRESULT, 'Revert'),
    COMMETHOD([], HRESULT, 'LockRegion',
              ( ['in'], _ULARGE_INTEGER, 'libOffset' ),
              ( ['in'], _ULARGE_INTEGER, 'cb' ),
              ( ['in'], c_ulong, 'dwLockType' )),
    COMMETHOD([], HRESULT, 'UnlockRegion',
              ( ['in'], _ULARGE_INTEGER, 'libOffset' ),
              ( ['in'], _ULARGE_INTEGER, 'cb' ),
              ( ['in'], c_ulong, 'dwLockType' )),
    COMMETHOD([], HRESULT, 'Stat',
              ( ['out'], POINTER(tagSTATSTG), 'pstatstg' ),
              ( ['in'], c_ulong, 'grfStatFlag' )),
    COMMETHOD([], HRESULT, 'Clone',
              ( ['out'], POINTER(POINTER(IStream)), 'ppstm' )),
]
################################################################
## code template for IStream implementation
##class IStream_Impl(object):
##    def RemoteSeek(self, dlibMove, dwOrigin):
##        '-no docstring-'
##        #return plibNewPosition
##
##    def Stat(self, grfStatFlag):
##        '-no docstring-'
##        #return pstatstg
##
##    def UnlockRegion(self, libOffset, cb, dwLockType):
##        '-no docstring-'
##        #return 
##
##    def Clone(self):
##        '-no docstring-'
##        #return ppstm
##
##    def Revert(self):
##        '-no docstring-'
##        #return 
##
##    def RemoteCopyTo(self, pstm, cb):
##        '-no docstring-'
##        #return pcbRead, pcbWritten
##
##    def LockRegion(self, libOffset, cb, dwLockType):
##        '-no docstring-'
##        #return 
##
##    def Commit(self, grfCommitFlags):
##        '-no docstring-'
##        #return 
##
##    def SetSize(self, libNewSize):
##        '-no docstring-'
##        #return 
##


# values for enumeration '__MIDL___MIDL_itf_tom_0000_0000_0001'
tomFalse = 0
tomTrue = -1
tomUndefined = -9999999
tomToggle = -9999998
tomAutoColor = -9999997
tomDefault = -9999996
tomSuspend = -9999995
tomResume = -9999994
tomApplyNow = 0
tomApplyLater = 1
tomTrackParms = 2
tomCacheParms = 3
tomApplyTmp = 4
tomDisableSmartFont = 8
tomEnableSmartFont = 9
tomUsePoints = 10
tomUseTwips = 11
tomBackward = -1073741823
tomForward = 1073741823
tomMove = 0
tomExtend = 1
tomNoSelection = 0
tomSelectionIP = 1
tomSelectionNormal = 2
tomSelectionFrame = 3
tomSelectionColumn = 4
tomSelectionRow = 5
tomSelectionBlock = 6
tomSelectionInlineShape = 7
tomSelectionShape = 8
tomSelStartActive = 1
tomSelAtEOL = 2
tomSelOvertype = 4
tomSelActive = 8
tomSelReplace = 16
tomEnd = 0
tomStart = 32
tomCollapseEnd = 0
tomCollapseStart = 1
tomClientCoord = 256
tomAllowOffClient = 512
tomTransform = 1024
tomObjectArg = 2048
tomAtEnd = 4096
tomUseXamlRect = 8192
tomImageTypeMask = 224
tomInlineImage = 0
tomBackgroundImage = 32
tomWrapTextAround = 64
tomNoWrapSides = 96
tomNone = 0
tomSingle = 1
tomWords = 2
tomDouble = 3
tomDotted = 4
tomDash = 5
tomDashDot = 6
tomDashDotDot = 7
tomWave = 8
tomThick = 9
tomHair = 10
tomDoubleWave = 11
tomHeavyWave = 12
tomLongDash = 13
tomThickDash = 14
tomThickDashDot = 15
tomThickDashDotDot = 16
tomThickDotted = 17
tomThickLongDash = 18
tomLineSpaceSingle = 0
tomLineSpace1pt5 = 1
tomLineSpaceDouble = 2
tomLineSpaceAtLeast = 3
tomLineSpaceExactly = 4
tomLineSpaceMultiple = 5
tomLineSpacePercent = 6
tomAlignLeft = 0
tomAlignCenter = 1
tomAlignRight = 2
tomAlignJustify = 3
tomAlignDecimal = 3
tomAlignBar = 4
tomDefaultTab = 5
tomAlignInterWord = 3
tomAlignNewspaper = 4
tomAlignInterLetter = 5
tomAlignScaled = 6
tomSpaces = 0
tomDots = 1
tomDashes = 2
tomLines = 3
tomThickLines = 4
tomEquals = 5
tomTabBack = -3
tomTabNext = -2
tomTabHere = -1
tomListNone = 0
tomListBullet = 1
tomListNumberAsArabic = 2
tomListNumberAsLCLetter = 3
tomListNumberAsUCLetter = 4
tomListNumberAsLCRoman = 5
tomListNumberAsUCRoman = 6
tomListNumberAsSequence = 7
tomListNumberedCircle = 8
tomListNumberedBlackCircleWingding = 9
tomListNumberedWhiteCircleWingding = 10
tomListNumberedArabicWide = 11
tomListNumberedChS = 12
tomListNumberedChT = 13
tomListNumberedJpnChS = 14
tomListNumberedJpnKor = 15
tomListNumberedArabic1 = 16
tomListNumberedArabic2 = 17
tomListNumberedHebrew = 18
tomListNumberedThaiAlpha = 19
tomListNumberedThaiNum = 20
tomListNumberedHindiAlpha = 21
tomListNumberedHindiAlpha1 = 22
tomListNumberedHindiNum = 23
tomListParentheses = 65536
tomListPeriod = 131072
tomListPlain = 196608
tomListNoNumber = 262144
tomListMinus = 524288
tomIgnoreNumberStyle = 16777216
tomParaStyleNormal = -1
tomParaStyleHeading1 = -2
tomParaStyleHeading2 = -3
tomParaStyleHeading3 = -4
tomParaStyleHeading4 = -5
tomParaStyleHeading5 = -6
tomParaStyleHeading6 = -7
tomParaStyleHeading7 = -8
tomParaStyleHeading8 = -9
tomParaStyleHeading9 = -10
tomCharacter = 1
tomWord = 2
tomSentence = 3
tomParagraph = 4
tomLine = 5
tomStory = 6
tomScreen = 7
tomSection = 8
tomTableColumn = 9
tomColumn = 9
tomRow = 10
tomWindow = 11
tomCell = 12
tomCharFormat = 13
tomParaFormat = 14
tomTable = 15
tomObject = 16
tomPage = 17
tomHardParagraph = 18
tomCluster = 19
tomInlineObject = 20
tomInlineObjectArg = 21
tomLeafLine = 22
tomLayoutColumn = 23
tomProcessId = 1073741825
tomMatchWord = 2
tomMatchCase = 4
tomMatchPattern = 8
tomUnknownStory = 0
tomMainTextStory = 1
tomFootnotesStory = 2
tomEndnotesStory = 3
tomCommentsStory = 4
tomTextFrameStory = 5
tomEvenPagesHeaderStory = 6
tomPrimaryHeaderStory = 7
tomEvenPagesFooterStory = 8
tomPrimaryFooterStory = 9
tomFirstPageHeaderStory = 10
tomFirstPageFooterStory = 11
tomScratchStory = 127
tomFindStory = 128
tomReplaceStory = 129
tomStoryInactive = 0
tomStoryActiveDisplay = 1
tomStoryActiveUI = 2
tomStoryActiveDisplayUI = 3
tomNoAnimation = 0
tomLasVegasLights = 1
tomBlinkingBackground = 2
tomSparkleText = 3
tomMarchingBlackAnts = 4
tomMarchingRedAnts = 5
tomShimmer = 6
tomWipeDown = 7
tomWipeRight = 8
tomAnimationMax = 8
tomLowerCase = 0
tomUpperCase = 1
tomTitleCase = 2
tomSentenceCase = 4
tomToggleCase = 5
tomReadOnly = 256
tomShareDenyRead = 512
tomShareDenyWrite = 1024
tomPasteFile = 4096
tomCreateNew = 16
tomCreateAlways = 32
tomOpenExisting = 48
tomOpenAlways = 64
tomTruncateExisting = 80
tomRTF = 1
tomText = 2
tomHTML = 3
tomWordDocument = 4
tomBold = -2147483647
tomItalic = -2147483646
tomUnderline = -2147483644
tomStrikeout = -2147483640
tomProtected = -2147483632
tomLink = -2147483616
tomSmallCaps = -2147483584
tomAllCaps = -2147483520
tomHidden = -2147483392
tomOutline = -2147483136
tomShadow = -2147482624
tomEmboss = -2147481600
tomImprint = -2147479552
tomDisabled = -2147475456
tomRevised = -2147467264
tomSubscriptCF = -2147418112
tomSuperscriptCF = -2147352576
tomFontBound = -2146435072
tomLinkProtected = -2139095040
tomInlineObjectStart = -2130706432
tomExtendedChar = -2113929216
tomAutoBackColor = -2080374784
tomMathZoneNoBuildUp = -2013265920
tomMathZone = -1879048192
tomMathZoneOrdinary = -1610612736
tomAutoTextColor = -1073741824
tomMathZoneDisplay = 262144
tomParaEffectRTL = 1
tomParaEffectKeep = 2
tomParaEffectKeepNext = 4
tomParaEffectPageBreakBefore = 8
tomParaEffectNoLineNumber = 16
tomParaEffectNoWidowControl = 32
tomParaEffectDoNotHyphen = 64
tomParaEffectSideBySide = 128
tomParaEffectCollapsed = 256
tomParaEffectOutlineLevel = 512
tomParaEffectBox = 1024
tomParaEffectTableRowDelimiter = 4096
tomParaEffectTable = 16384
tomModWidthPairs = 1
tomModWidthSpace = 2
tomAutoSpaceAlpha = 4
tomAutoSpaceNumeric = 8
tomAutoSpaceParens = 16
tomEmbeddedFont = 32
tomDoublestrike = 64
tomOverlapping = 128
tomNormalCaret = 0
tomKoreanBlockCaret = 1
tomNullCaret = 2
tomIncludeInset = 1
tomUnicodeBiDi = 1
tomMathCFCheck = 4
tomUnlink = 8
tomUnhide = 16
tomCheckTextLimit = 32
tomDontSelectText = 64
tomIgnoreCurrentFont = 0
tomMatchCharRep = 1
tomMatchFontSignature = 2
tomMatchAscii = 4
tomGetHeightOnly = 8
tomMatchMathFont = 16
tomCharset = -2147483648
tomCharRepFromLcid = 1073741824
tomAnsi = 0
tomEastEurope = 1
tomCyrillic = 2
tomGreek = 3
tomTurkish = 4
tomHebrew = 5
tomArabic = 6
tomBaltic = 7
tomVietnamese = 8
tomDefaultCharRep = 9
tomSymbol = 10
tomThai = 11
tomShiftJIS = 12
tomGB2312 = 13
tomHangul = 14
tomBIG5 = 15
tomPC437 = 16
tomOEM = 17
tomMac = 18
tomArmenian = 19
tomSyriac = 20
tomThaana = 21
tomDevanagari = 22
tomBengali = 23
tomGurmukhi = 24
tomGujarati = 25
tomOriya = 26
tomTamil = 27
tomTelugu = 28
tomKannada = 29
tomMalayalam = 30
tomSinhala = 31
tomLao = 32
tomTibetan = 33
tomMyanmar = 34
tomGeorgian = 35
tomJamo = 36
tomEthiopic = 37
tomCherokee = 38
tomAboriginal = 39
tomOgham = 40
tomRunic = 41
tomKhmer = 42
tomMongolian = 43
tomBraille = 44
tomYi = 45
tomLimbu = 46
tomTaiLe = 47
tomNewTaiLue = 48
tomSylotiNagri = 49
tomKharoshthi = 50
tomKayahli = 51
tomUsymbol = 52
tomEmoji = 53
tomGlagolitic = 54
tomLisu = 55
tomVai = 56
tomNKo = 57
tomOsmanya = 58
tomPhagsPa = 59
tomGothic = 60
tomDeseret = 61
tomTifinagh = 62
tomCharRepMax = 63
tomRE10Mode = 1
tomUseAtFont = 2
tomTextFlowMask = 12
tomTextFlowES = 0
tomTextFlowSW = 4
tomTextFlowWN = 8
tomTextFlowNE = 12
tomNoIME = 524288
tomSelfIME = 262144
tomNoUpScroll = 65536
tomNoVpScroll = 262144
tomNoLink = 0
tomClientLink = 1
tomFriendlyLinkName = 2
tomFriendlyLinkAddress = 3
tomAutoLinkURL = 4
tomAutoLinkEmail = 5
tomAutoLinkPhone = 6
tomAutoLinkPath = 7
tomCompressNone = 0
tomCompressPunctuation = 1
tomCompressPunctuationAndKana = 2
tomCompressMax = 2
tomUnderlinePositionAuto = 0
tomUnderlinePositionBelow = 1
tomUnderlinePositionAbove = 2
tomUnderlinePositionMax = 2
tomFontAlignmentAuto = 0
tomFontAlignmentTop = 1
tomFontAlignmentBaseline = 2
tomFontAlignmentBottom = 3
tomFontAlignmentCenter = 4
tomFontAlignmentMax = 4
tomRubyBelow = 128
tomRubyAlignCenter = 0
tomRubyAlign010 = 1
tomRubyAlign121 = 2
tomRubyAlignLeft = 3
tomRubyAlignRight = 4
tomLimitsDefault = 0
tomLimitsUnderOver = 1
tomLimitsSubSup = 2
tomUpperLimitAsSuperScript = 3
tomLimitsOpposite = 4
tomShowLLimPlaceHldr = 8
tomShowULimPlaceHldr = 16
tomDontGrowWithContent = 64
tomGrowWithContent = 128
tomSubSupAlign = 1
tomLimitAlignMask = 3
tomLimitAlignCenter = 0
tomLimitAlignLeft = 1
tomLimitAlignRight = 2
tomShowDegPlaceHldr = 8
tomAlignDefault = 0
tomAlignMatchAscentDescent = 2
tomMathVariant = 32
tomStyleDefault = 0
tomStyleScriptScriptCramped = 1
tomStyleScriptScript = 2
tomStyleScriptCramped = 3
tomStyleScript = 4
tomStyleTextCramped = 5
tomStyleText = 6
tomStyleDisplayCramped = 7
tomStyleDisplay = 8
tomMathRelSize = 64
tomDecDecSize = 254
tomDecSize = 255
tomIncSize = 65
tomIncIncSize = 66
tomGravityUI = 0
tomGravityBack = 1
tomGravityFore = 2
tomGravityIn = 3
tomGravityOut = 4
tomGravityBackward = 536870912
tomGravityForward = 1073741824
tomTeX = 1
tomNeedTermOp = 2
tomMathAlphabetics = 4
tomMathSingleChar = 8
tomPlain = 16
tomHaveDelimiter = 32
tomUseOperandPrec = 64
tomMathCollapseSel = 128
tomMathAutoCorrect = 256
tomMathBuildUpArgOrZone = 512
tomMathBuildUpRecurse = 1024
tomMathBuildDownOutermost = 2048
tomChemicalFormula = 4096
tomMathBuildDown = 8192
tomMathApplyTemplate = 16384
tomMathRemoveOutermost = 32768
tomMathChangeMask = 2031616
tomMathInsRowBefore = 65536
tomMathInsRowAfter = 131072
tomMathInsColBefore = 196608
tomMathInsColAfter = 262144
tomMathDeleteRow = 327680
tomMathDeleteCol = 393216
tomMathDeleteArg = 458752
tomMathDeleteArg1 = 524288
tomMathDeleteArg2 = 589824
tomMathMakeFracLinear = 655360
tomMathMakeFracStacked = 720896
tomMathMakeFracSlashed = 786432
tomMathMakeLeftSubSup = 851968
tomMathMakeSubSup = 917504
tomMathBackspace = 1048576
tomMathEnter = 1114112
tomMathShiftTab = 1179648
tomMathTab = 1245184
tomMathAlignBreakLeft = 1310720
tomMathAlignBreakCenter = 1376256
tomMathAlignBreakRight = 1441792
tomMathSubscript = 1507328
tomMathSuperscript = 1572864
tomMathArabicAlphabetics = 8388608
tomMathAutoCorrectOpPairs = 16777216
tomMathAutoCorrectExt = 33554432
tomShowEmptyArgPlaceholders = 67108864
tomMathAutoComplete = 134217728
tomMathRichEdit = 1073741824
tomSpecialChar = -2147483648
tomAdjustCRLF = 1
tomUseCRLF = 2
tomTextize = 4
tomAllowFinalEOP = 8
tomFoldMathAlpha = 16
tomNoHidden = 32
tomIncludeNumbering = 64
tomTranslateTableCell = 128
tomNoMathZoneBrackets = 256
tomConvertMathChar = 512
tomNoUCGreekItalic = 1024
tomAllowMathBold = 2048
tomLanguageTag = 4096
tomConvertRTF = 8192
tomApplyRtfDocProps = 16384
tomGetTextForSpell = 32768
tomConvertMathML = 65536
tomGetUtf16 = 131072
tomConvertLinearFormat = 262144
tomConvertOMML = 524288
tomConvertRuby = 1048576
tomConvertCRtoLF = 16777216
tomPhantomShow = 1
tomPhantomZeroWidth = 2
tomPhantomZeroAscent = 4
tomPhantomZeroDescent = 8
tomPhantomTransparent = 16
tomPhantomASmash = 5
tomPhantomDSmash = 9
tomPhantomHSmash = 3
tomPhantomSmash = 13
tomPhantomHorz = 12
tomPhantomVert = 2
tomBoxHideTop = 1
tomBoxHideBottom = 2
tomBoxHideLeft = 4
tomBoxHideRight = 8
tomBoxStrikeH = 16
tomBoxStrikeV = 32
tomBoxStrikeTLBR = 64
tomBoxStrikeBLTR = 128
tomRoundedBoxDashStyleMask = 7
tomRoundedBoxHideBorder = 8
tomRoundedBoxCapStyleMask = 48
tomRoundedBoxNullRadius = 64
tomRoundedBoxCompact = 128
tomBoxAlignCenter = 1
tomSpaceMask = 28
tomSpaceDefault = 0
tomSpaceUnary = 4
tomSpaceBinary = 8
tomSpaceRelational = 12
tomSpaceSkip = 16
tomSpaceOrd = 20
tomSpaceDifferential = 24
tomSizeText = 32
tomSizeScript = 64
tomSizeScriptScript = 96
tomNoBreak = 128
tomTransparentForPositioning = 256
tomTransparentForSpacing = 512
tomStretchCharBelow = 0
tomStretchCharAbove = 1
tomStretchBaseBelow = 2
tomStretchBaseAbove = 3
tomMatrixAlignMask = 3
tomMatrixAlignCenter = 0
tomMatrixAlignTopRow = 1
tomMatrixAlignBottomRow = 3
tomShowMatPlaceHldr = 8
tomEqArrayLayoutWidth = 1
tomEqArrayAlignMask = 12
tomEqArrayAlignCenter = 0
tomEqArrayAlignTopRow = 4
tomEqArrayAlignBottomRow = 12
tomMathManualBreakMask = 127
tomMathBreakLeft = 125
tomMathBreakCenter = 126
tomMathBreakRight = 127
tomMathEqAlign = 128
tomMathArgShadingStart = 593
tomMathArgShadingEnd = 594
tomMathObjShadingStart = 595
tomMathObjShadingEnd = 596
tomFunctionTypeNone = 0
tomFunctionTypeTakesArg = 1
tomFunctionTypeTakesLim = 2
tomFunctionTypeTakesLim2 = 3
tomFunctionTypeIsLim = 4
tomMathParaAlignDefault = 0
tomMathParaAlignCenterGroup = 1
tomMathParaAlignCenter = 2
tomMathParaAlignLeft = 3
tomMathParaAlignRight = 4
tomMathDispAlignMask = 3
tomMathDispAlignCenterGroup = 0
tomMathDispAlignCenter = 1
tomMathDispAlignLeft = 2
tomMathDispAlignRight = 3
tomMathDispIntUnderOver = 4
tomMathDispFracTeX = 8
tomMathDispNaryGrow = 16
tomMathDocEmptyArgMask = 96
tomMathDocEmptyArgAuto = 0
tomMathDocEmptyArgAlways = 32
tomMathDocEmptyArgNever = 64
tomMathDocSbSpOpUnchanged = 128
tomMathDocDiffMask = 768
tomMathDocDiffDefault = 0
tomMathDocDiffUpright = 256
tomMathDocDiffItalic = 512
tomMathDocDiffOpenItalic = 768
tomMathDispNarySubSup = 1024
tomMathDispDef = 2048
tomMathEnableRtl = 4096
tomMathBrkBinMask = 196608
tomMathBrkBinBefore = 0
tomMathBrkBinAfter = 65536
tomMathBrkBinDup = 131072
tomMathBrkBinSubMask = 786432
tomMathBrkBinSubMM = 0
tomMathBrkBinSubPM = 262144
tomMathBrkBinSubMP = 524288
tomSelRange = 597
tomHstring = 596
tomFontPropTeXStyle = 828
tomFontPropAlign = 829
tomFontStretch = 830
tomFontStyle = 831
tomFontStyleUpright = 0
tomFontStyleOblique = 1
tomFontStyleItalic = 2
tomFontStretchDefault = 0
tomFontStretchUltraCondensed = 1
tomFontStretchExtraCondensed = 2
tomFontStretchCondensed = 3
tomFontStretchSemiCondensed = 4
tomFontStretchNormal = 5
tomFontStretchSemiExpanded = 6
tomFontStretchExpanded = 7
tomFontStretchExtraExpanded = 8
tomFontStretchUltraExpanded = 9
tomFontWeightDefault = 0
tomFontWeightThin = 100
tomFontWeightExtraLight = 200
tomFontWeightLight = 300
tomFontWeightNormal = 400
tomFontWeightRegular = 400
tomFontWeightMedium = 500
tomFontWeightSemiBold = 600
tomFontWeightBold = 700
tomFontWeightExtraBold = 800
tomFontWeightBlack = 900
tomFontWeightHeavy = 900
tomFontWeightExtraBlack = 950
tomFontWeightMax = 950
tomParaPropMathAlign = 1079
tomDocMathBuild = 128
tomMathLMargin = 129
tomMathRMargin = 130
tomMathWrapIndent = 131
tomMathWrapRight = 132
tomMathPostSpace = 134
tomMathPreSpace = 133
tomMathInterSpace = 135
tomMathIntraSpace = 136
tomCanCopy = 137
tomCanRedo = 138
tomCanUndo = 139
tomUndoLimit = 140
tomDocAutoLink = 141
tomEllipsisMode = 142
tomEllipsisState = 143
tomMathZoneSurround = 144
tomUnderlineTrailSpace = 145
tomAlignWithTrailSpace = 146
tomIgnoreTrailSpacing = 147
tomStoryLength = 2624
tomEllipsisNone = 0
tomEllipsisEnd = 1
tomEllipsisWord = 3
tomEllipsisPresent = 1
tomVTopCell = 1
tomVLowCell = 2
tomHStartCell = 4
tomHContCell = 8
tomRowUpdate = 1
tomRowApplyDefault = 0
tomCellStructureChangeOnly = 1
tomRowHeightActual = 2059
tomImageFlipH = 1
tomImageFlipV = 2
tomImageRotate0 = 0
tomImageRotate90 = 1
tomImageRotate180 = 2
tomImageRotate270 = 3
__MIDL___MIDL_itf_tom_0000_0000_0001 = c_int # enum
tomConstants = __MIDL___MIDL_itf_tom_0000_0000_0001
class ITextRange(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{8CC497C2-A1DF-11CE-8098-00AA0047BE5D}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class ITextSelection(ITextRange):
    _case_insensitive_ = True
    _iid_ = GUID('{8CC497C1-A1DF-11CE-8098-00AA0047BE5D}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class ITextFont(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{8CC497C3-A1DF-11CE-8098-00AA0047BE5D}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class ITextPara(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{8CC497C4-A1DF-11CE-8098-00AA0047BE5D}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextRange._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Text',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Text',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(513), 'propget'], HRESULT, 'Char',
              ( ['retval', 'out'], POINTER(c_int), 'pChar' )),
    COMMETHOD([dispid(513), 'propput'], HRESULT, 'Char',
              ( ['in'], c_int, 'pChar' )),
    COMMETHOD([dispid(514), 'propget'], HRESULT, 'Duplicate',
              ( ['retval', 'out'], POINTER(POINTER(ITextRange)), 'ppRange' )),
    COMMETHOD([dispid(515), 'propget'], HRESULT, 'FormattedText',
              ( ['retval', 'out'], POINTER(POINTER(ITextRange)), 'ppRange' )),
    COMMETHOD([dispid(515), 'propput'], HRESULT, 'FormattedText',
              ( ['in'], POINTER(ITextRange), 'ppRange' )),
    COMMETHOD([dispid(516), 'propget'], HRESULT, 'Start',
              ( ['retval', 'out'], POINTER(c_int), 'pcpFirst' )),
    COMMETHOD([dispid(516), 'propput'], HRESULT, 'Start',
              ( ['in'], c_int, 'pcpFirst' )),
    COMMETHOD([dispid(517), 'propget'], HRESULT, 'End',
              ( ['retval', 'out'], POINTER(c_int), 'pcpLim' )),
    COMMETHOD([dispid(517), 'propput'], HRESULT, 'End',
              ( ['in'], c_int, 'pcpLim' )),
    COMMETHOD([dispid(518), 'propget'], HRESULT, 'Font',
              ( ['retval', 'out'], POINTER(POINTER(ITextFont)), 'ppFont' )),
    COMMETHOD([dispid(518), 'propput'], HRESULT, 'Font',
              ( ['in'], POINTER(ITextFont), 'ppFont' )),
    COMMETHOD([dispid(519), 'propget'], HRESULT, 'Para',
              ( ['retval', 'out'], POINTER(POINTER(ITextPara)), 'ppPara' )),
    COMMETHOD([dispid(519), 'propput'], HRESULT, 'Para',
              ( ['in'], POINTER(ITextPara), 'ppPara' )),
    COMMETHOD([dispid(520), 'propget'], HRESULT, 'StoryLength',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(521), 'propget'], HRESULT, 'StoryType',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(528)], HRESULT, 'Collapse',
              ( ['in'], c_int, 'bStart' )),
    COMMETHOD([dispid(529)], HRESULT, 'Expand',
              ( ['in'], c_int, 'Unit' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(530)], HRESULT, 'GetIndex',
              ( ['in'], c_int, 'Unit' ),
              ( ['retval', 'out'], POINTER(c_int), 'pIndex' )),
    COMMETHOD([dispid(531)], HRESULT, 'SetIndex',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Index' ),
              ( ['in'], c_int, 'Extend' )),
    COMMETHOD([dispid(532)], HRESULT, 'SetRange',
              ( ['in'], c_int, 'cpAnchor' ),
              ( ['in'], c_int, 'cpActive' )),
    COMMETHOD([dispid(533)], HRESULT, 'InRange',
              ( ['in'], POINTER(ITextRange), 'pRange' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(534)], HRESULT, 'InStory',
              ( ['in'], POINTER(ITextRange), 'pRange' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(535)], HRESULT, 'IsEqual',
              ( ['in'], POINTER(ITextRange), 'pRange' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(536)], HRESULT, 'Select'),
    COMMETHOD([dispid(537)], HRESULT, 'StartOf',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(544)], HRESULT, 'EndOf',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(545)], HRESULT, 'Move',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(546)], HRESULT, 'MoveStart',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(547)], HRESULT, 'MoveEnd',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(548)], HRESULT, 'MoveWhile',
              ( ['in'], POINTER(VARIANT), 'Cset' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(549)], HRESULT, 'MoveStartWhile',
              ( ['in'], POINTER(VARIANT), 'Cset' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(550)], HRESULT, 'MoveEndWhile',
              ( ['in'], POINTER(VARIANT), 'Cset' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(551)], HRESULT, 'MoveUntil',
              ( ['in'], POINTER(VARIANT), 'Cset' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(552)], HRESULT, 'MoveStartUntil',
              ( ['in'], POINTER(VARIANT), 'Cset' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(553)], HRESULT, 'MoveEndUntil',
              ( ['in'], POINTER(VARIANT), 'Cset' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(560)], HRESULT, 'FindText',
              ( ['in'], BSTR, 'bstr' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Flags' ),
              ( ['retval', 'out'], POINTER(c_int), 'pLength' )),
    COMMETHOD([dispid(561)], HRESULT, 'FindTextStart',
              ( ['in'], BSTR, 'bstr' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Flags' ),
              ( ['retval', 'out'], POINTER(c_int), 'pLength' )),
    COMMETHOD([dispid(562)], HRESULT, 'FindTextEnd',
              ( ['in'], BSTR, 'bstr' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Flags' ),
              ( ['retval', 'out'], POINTER(c_int), 'pLength' )),
    COMMETHOD([dispid(563)], HRESULT, 'Delete',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(564)], HRESULT, 'Cut',
              ( ['out'], POINTER(VARIANT), 'pVar' )),
    COMMETHOD([dispid(565)], HRESULT, 'Copy',
              ( ['out'], POINTER(VARIANT), 'pVar' )),
    COMMETHOD([dispid(566)], HRESULT, 'Paste',
              ( ['in'], POINTER(VARIANT), 'pVar' ),
              ( ['in'], c_int, 'Format' )),
    COMMETHOD([dispid(567)], HRESULT, 'CanPaste',
              ( ['in'], POINTER(VARIANT), 'pVar' ),
              ( ['in'], c_int, 'Format' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(568)], HRESULT, 'CanEdit',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(569)], HRESULT, 'ChangeCase',
              ( ['in'], c_int, 'Type' )),
    COMMETHOD([dispid(576)], HRESULT, 'GetPoint',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'px' ),
              ( ['out'], POINTER(c_int), 'py' )),
    COMMETHOD([dispid(577)], HRESULT, 'SetPoint',
              ( ['in'], c_int, 'x' ),
              ( ['in'], c_int, 'y' ),
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Extend' )),
    COMMETHOD([dispid(578)], HRESULT, 'ScrollIntoView',
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(579)], HRESULT, 'GetEmbeddedObject',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppObject' )),
]
################################################################
## code template for ITextRange implementation
##class ITextRange_Impl(object):
##    def MoveEndWhile(self, Cset, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def Cut(self):
##        '-no docstring-'
##        #return pVar
##
##    def _get(self):
##        '-no docstring-'
##        #return pcpLim
##    def _set(self, pcpLim):
##        '-no docstring-'
##    End = property(_get, _set, doc = _set.__doc__)
##
##    def MoveStartWhile(self, Cset, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def FindText(self, bstr, Count, Flags):
##        '-no docstring-'
##        #return pLength
##
##    def SetPoint(self, x, y, Type, Extend):
##        '-no docstring-'
##        #return 
##
##    def GetEmbeddedObject(self):
##        '-no docstring-'
##        #return ppObject
##
##    def FindTextStart(self, bstr, Count, Flags):
##        '-no docstring-'
##        #return pLength
##
##    def _get(self):
##        '-no docstring-'
##        #return pcpFirst
##    def _set(self, pcpFirst):
##        '-no docstring-'
##    Start = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Duplicate(self):
##        '-no docstring-'
##        #return ppRange
##
##    def MoveUntil(self, Cset, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def IsEqual(self, pRange):
##        '-no docstring-'
##        #return pValue
##
##    def MoveWhile(self, Cset, Count):
##        '-no docstring-'
##        #return pDelta
##
##    @property
##    def StoryLength(self):
##        '-no docstring-'
##        #return pCount
##
##    def CanEdit(self):
##        '-no docstring-'
##        #return pValue
##
##    def MoveStart(self, Unit, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def Collapse(self, bStart):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return ppPara
##    def _set(self, ppPara):
##        '-no docstring-'
##    Para = property(_get, _set, doc = _set.__doc__)
##
##    def SetIndex(self, Unit, Index, Extend):
##        '-no docstring-'
##        #return 
##
##    def EndOf(self, Unit, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    def GetPoint(self, Type):
##        '-no docstring-'
##        #return px, py
##
##    def CanPaste(self, pVar, Format):
##        '-no docstring-'
##        #return pValue
##
##    def InStory(self, pRange):
##        '-no docstring-'
##        #return pValue
##
##    def Copy(self):
##        '-no docstring-'
##        #return pVar
##
##    def Paste(self, pVar, Format):
##        '-no docstring-'
##        #return 
##
##    def Expand(self, Unit):
##        '-no docstring-'
##        #return pDelta
##
##    def GetIndex(self, Unit):
##        '-no docstring-'
##        #return pIndex
##
##    def MoveEnd(self, Unit, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def _get(self):
##        '-no docstring-'
##        #return ppRange
##    def _set(self, ppRange):
##        '-no docstring-'
##    FormattedText = property(_get, _set, doc = _set.__doc__)
##
##    def ScrollIntoView(self, Value):
##        '-no docstring-'
##        #return 
##
##    def FindTextEnd(self, bstr, Count, Flags):
##        '-no docstring-'
##        #return pLength
##
##    def ChangeCase(self, Type):
##        '-no docstring-'
##        #return 
##
##    def InRange(self, pRange):
##        '-no docstring-'
##        #return pValue
##
##    def Delete(self, Unit, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    Text = property(_get, _set, doc = _set.__doc__)
##
##    def Move(self, Unit, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def MoveEndUntil(self, Cset, Count):
##        '-no docstring-'
##        #return pDelta
##
##    def _get(self):
##        '-no docstring-'
##        #return pChar
##    def _set(self, pChar):
##        '-no docstring-'
##    Char = property(_get, _set, doc = _set.__doc__)
##
##    def MoveStartUntil(self, Cset, Count):
##        '-no docstring-'
##        #return pDelta
##
##    @property
##    def StoryType(self):
##        '-no docstring-'
##        #return pValue
##
##    def StartOf(self, Unit, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    def _get(self):
##        '-no docstring-'
##        #return ppFont
##    def _set(self, ppFont):
##        '-no docstring-'
##    Font = property(_get, _set, doc = _set.__doc__)
##
##    def Select(self):
##        '-no docstring-'
##        #return 
##
##    def SetRange(self, cpAnchor, cpActive):
##        '-no docstring-'
##        #return 
##

ITextSelection._methods_ = [
    COMMETHOD([dispid(257), 'propget'], HRESULT, 'Flags',
              ( ['retval', 'out'], POINTER(c_int), 'pFlags' )),
    COMMETHOD([dispid(257), 'propput'], HRESULT, 'Flags',
              ( ['in'], c_int, 'pFlags' )),
    COMMETHOD([dispid(258), 'propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(c_int), 'pType' )),
    COMMETHOD([dispid(259)], HRESULT, 'MoveLeft',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(260)], HRESULT, 'MoveRight',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(261)], HRESULT, 'MoveUp',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(262)], HRESULT, 'MoveDown',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(263)], HRESULT, 'HomeKey',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(264)], HRESULT, 'EndKey',
              ( ['in'], c_int, 'Unit' ),
              ( ['in'], c_int, 'Extend' ),
              ( ['retval', 'out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(265)], HRESULT, 'TypeText',
              ( ['in'], BSTR, 'bstr' )),
]
################################################################
## code template for ITextSelection implementation
##class ITextSelection_Impl(object):
##    def MoveLeft(self, Unit, Count, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    def EndKey(self, Unit, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    def MoveDown(self, Unit, Count, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    def HomeKey(self, Unit, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    def _get(self):
##        '-no docstring-'
##        #return pFlags
##    def _set(self, pFlags):
##        '-no docstring-'
##    Flags = property(_get, _set, doc = _set.__doc__)
##
##    def MoveRight(self, Unit, Count, Extend):
##        '-no docstring-'
##        #return pDelta
##
##    @property
##    def Type(self):
##        '-no docstring-'
##        #return pType
##
##    def TypeText(self, bstr):
##        '-no docstring-'
##        #return 
##
##    def MoveUp(self, Unit, Count, Extend):
##        '-no docstring-'
##        #return pDelta
##

class ITextDisplays(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5F2-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextDisplays._methods_ = [
]
################################################################
## code template for ITextDisplays implementation
##class ITextDisplays_Impl(object):

class Library(object):
    name = u'tom'
    _reg_typelib_ = ('{8CC497C9-A1DF-11CE-8098-00AA0047BE5D}', 1, 0)

class ITextStoryRanges(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{8CC497C5-A1DF-11CE-8098-00AA0047BE5D}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class ITextStoryRanges2(ITextStoryRanges):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E5-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextStoryRanges._methods_ = [
    COMMETHOD([dispid(-4), 'restricted'], HRESULT, '_NewEnum',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppunkEnum' )),
    COMMETHOD([dispid(0)], HRESULT, 'Item',
              ( ['in'], c_int, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange)), 'ppRange' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'Count',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
]
################################################################
## code template for ITextStoryRanges implementation
##class ITextStoryRanges_Impl(object):
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pCount
##
##    def Item(self, Index):
##        '-no docstring-'
##        #return ppRange
##
##    def _NewEnum(self):
##        '-no docstring-'
##        #return ppunkEnum
##

class ITextRange2(ITextSelection):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextStoryRanges2._methods_ = [
    COMMETHOD([dispid(3)], HRESULT, 'Item2',
              ( ['in'], c_int, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
]
################################################################
## code template for ITextStoryRanges2 implementation
##class ITextStoryRanges2_Impl(object):
##    def Item2(self, Index):
##        '-no docstring-'
##        #return ppRange
##


# values for enumeration '__MIDL___MIDL_itf_tom_0000_0000_0002'
tomSimpleText = 0
tomRuby = 1
tomHorzVert = 2
tomWarichu = 3
tomHorzHorz = 4
tomEnclose = 5
tomEq = 9
tomMath = 10
tomAccent = 10
tomBox = 11
tomBoxedFormula = 12
tomBrackets = 13
tomBracketsWithSeps = 14
tomEquationArray = 15
tomFraction = 16
tomFunctionApply = 17
tomLeftSubSup = 18
tomLowerLimit = 19
tomMatrix = 20
tomNary = 21
tomOpChar = 22
tomOverbar = 23
tomPhantom = 24
tomRadical = 25
tomSlashedFraction = 26
tomStack = 27
tomStretchStack = 28
tomSubscript = 29
tomSubSup = 30
tomSuperscript = 31
tomUnderbar = 32
tomUpperLimit = 33
tomObjectMax = 33
__MIDL___MIDL_itf_tom_0000_0000_0002 = c_int # enum
OBJECTTYPE = __MIDL___MIDL_itf_tom_0000_0000_0002

# values for enumeration '__MIDL___MIDL_itf_tom_0000_0000_0003'
MBOLD = 16
MITAL = 32
MGREEK = 64
MROMN = 0
MSCRP = 1
MFRAK = 2
MOPEN = 3
MSANS = 4
MMONO = 5
MMATH = 6
MISOL = 7
MINIT = 8
MTAIL = 9
MSTRCH = 10
MLOOP = 11
MOPENA = 12
__MIDL___MIDL_itf_tom_0000_0000_0003 = c_int # enum
MANCODE = __MIDL___MIDL_itf_tom_0000_0000_0003
class ITextFont2(ITextFont):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextFont._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Duplicate',
              ( ['retval', 'out'], POINTER(POINTER(ITextFont)), 'ppFont' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Duplicate',
              ( ['in'], POINTER(ITextFont), 'ppFont' )),
    COMMETHOD([dispid(769)], HRESULT, 'CanChange',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(770)], HRESULT, 'IsEqual',
              ( ['in'], POINTER(ITextFont), 'pFont' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(771)], HRESULT, 'Reset',
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(772), 'propget'], HRESULT, 'Style',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(772), 'propput'], HRESULT, 'Style',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(773), 'propget'], HRESULT, 'AllCaps',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(773), 'propput'], HRESULT, 'AllCaps',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(774), 'propget'], HRESULT, 'Animation',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(774), 'propput'], HRESULT, 'Animation',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(775), 'propget'], HRESULT, 'BackColor',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(775), 'propput'], HRESULT, 'BackColor',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(776), 'propget'], HRESULT, 'Bold',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(776), 'propput'], HRESULT, 'Bold',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(777), 'propget'], HRESULT, 'Emboss',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(777), 'propput'], HRESULT, 'Emboss',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(784), 'propget'], HRESULT, 'ForeColor',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(784), 'propput'], HRESULT, 'ForeColor',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(785), 'propget'], HRESULT, 'Hidden',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(785), 'propput'], HRESULT, 'Hidden',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(786), 'propget'], HRESULT, 'Engrave',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(786), 'propput'], HRESULT, 'Engrave',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(787), 'propget'], HRESULT, 'Italic',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(787), 'propput'], HRESULT, 'Italic',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(788), 'propget'], HRESULT, 'Kerning',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(788), 'propput'], HRESULT, 'Kerning',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(789), 'propget'], HRESULT, 'LanguageID',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(789), 'propput'], HRESULT, 'LanguageID',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(790), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(790), 'propput'], HRESULT, 'Name',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(791), 'propget'], HRESULT, 'Outline',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(791), 'propput'], HRESULT, 'Outline',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(792), 'propget'], HRESULT, 'Position',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(792), 'propput'], HRESULT, 'Position',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(793), 'propget'], HRESULT, 'Protected',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(793), 'propput'], HRESULT, 'Protected',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(800), 'propget'], HRESULT, 'Shadow',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(800), 'propput'], HRESULT, 'Shadow',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(801), 'propget'], HRESULT, 'Size',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(801), 'propput'], HRESULT, 'Size',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(802), 'propget'], HRESULT, 'SmallCaps',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(802), 'propput'], HRESULT, 'SmallCaps',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(803), 'propget'], HRESULT, 'Spacing',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(803), 'propput'], HRESULT, 'Spacing',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(804), 'propget'], HRESULT, 'StrikeThrough',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(804), 'propput'], HRESULT, 'StrikeThrough',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(805), 'propget'], HRESULT, 'Subscript',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(805), 'propput'], HRESULT, 'Subscript',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(806), 'propget'], HRESULT, 'Superscript',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(806), 'propput'], HRESULT, 'Superscript',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(807), 'propget'], HRESULT, 'Underline',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(807), 'propput'], HRESULT, 'Underline',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(808), 'propget'], HRESULT, 'Weight',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(808), 'propput'], HRESULT, 'Weight',
              ( ['in'], c_int, 'pValue' )),
]
################################################################
## code template for ITextFont implementation
##class ITextFont_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Style = property(_get, _set, doc = _set.__doc__)
##
##    def CanChange(self):
##        '-no docstring-'
##        #return pValue
##
##    def IsEqual(self, pFont):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return ppFont
##    def _set(self, ppFont):
##        '-no docstring-'
##    Duplicate = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Animation = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Italic = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Engrave = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Hidden = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Subscript = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    Name = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Bold = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Spacing = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    AllCaps = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ForeColor = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    BackColor = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Shadow = property(_get, _set, doc = _set.__doc__)
##
##    def Reset(self, Value):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Outline = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    StrikeThrough = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    SmallCaps = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Protected = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Position = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Superscript = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Weight = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    LanguageID = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Kerning = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Emboss = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Underline = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Size = property(_get, _set, doc = _set.__doc__)
##

ITextFont2._methods_ = [
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'Count',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(809), 'propget'], HRESULT, 'AutoLigatures',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(809), 'propput'], HRESULT, 'AutoLigatures',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(810), 'propget'], HRESULT, 'AutospaceAlpha',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(810), 'propput'], HRESULT, 'AutospaceAlpha',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(811), 'propget'], HRESULT, 'AutospaceNumeric',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(811), 'propput'], HRESULT, 'AutospaceNumeric',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(812), 'propget'], HRESULT, 'AutospaceParens',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(812), 'propput'], HRESULT, 'AutospaceParens',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(813), 'propget'], HRESULT, 'CharRep',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(813), 'propput'], HRESULT, 'CharRep',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(814), 'propget'], HRESULT, 'CompressionMode',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(814), 'propput'], HRESULT, 'CompressionMode',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(815), 'propget'], HRESULT, 'Cookie',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(815), 'propput'], HRESULT, 'Cookie',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(816), 'propget'], HRESULT, 'DoubleStrike',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(816), 'propput'], HRESULT, 'DoubleStrike',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(817), 'propget'], HRESULT, 'Duplicate2',
              ( ['retval', 'out'], POINTER(POINTER(ITextFont2)), 'ppFont' )),
    COMMETHOD([dispid(817), 'propput'], HRESULT, 'Duplicate2',
              ( ['in'], POINTER(ITextFont2), 'ppFont' )),
    COMMETHOD([dispid(818), 'propget'], HRESULT, 'LinkType',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(819), 'propget'], HRESULT, 'MathZone',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(819), 'propput'], HRESULT, 'MathZone',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(820), 'propget'], HRESULT, 'ModWidthPairs',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(820), 'propput'], HRESULT, 'ModWidthPairs',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(821), 'propget'], HRESULT, 'ModWidthSpace',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(821), 'propput'], HRESULT, 'ModWidthSpace',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(822), 'propget'], HRESULT, 'OldNumbers',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(822), 'propput'], HRESULT, 'OldNumbers',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(823), 'propget'], HRESULT, 'Overlapping',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(823), 'propput'], HRESULT, 'Overlapping',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(824), 'propget'], HRESULT, 'PositionSubSuper',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(824), 'propput'], HRESULT, 'PositionSubSuper',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(825), 'propget'], HRESULT, 'Scaling',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(825), 'propput'], HRESULT, 'Scaling',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(826), 'propget'], HRESULT, 'SpaceExtension',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(826), 'propput'], HRESULT, 'SpaceExtension',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(827), 'propget'], HRESULT, 'UnderlinePositionMode',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(827), 'propput'], HRESULT, 'UnderlinePositionMode',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(832)], HRESULT, 'GetEffects',
              ( ['out'], POINTER(c_int), 'pValue' ),
              ( ['out'], POINTER(c_int), 'pMask' )),
    COMMETHOD([dispid(833)], HRESULT, 'GetEffects2',
              ( ['out'], POINTER(c_int), 'pValue' ),
              ( ['out'], POINTER(c_int), 'pMask' )),
    COMMETHOD([dispid(834)], HRESULT, 'GetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(835)], HRESULT, 'GetPropertyInfo',
              ( ['in'], c_int, 'Index' ),
              ( ['out'], POINTER(c_int), 'pType' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(836)], HRESULT, 'IsEqual2',
              ( ['in'], POINTER(ITextFont2), 'pFont' ),
              ( ['retval', 'out'], POINTER(c_int), 'pB' )),
    COMMETHOD([dispid(837)], HRESULT, 'SetEffects',
              ( ['in'], c_int, 'Value' ),
              ( ['in'], c_int, 'Mask' )),
    COMMETHOD([dispid(838)], HRESULT, 'SetEffects2',
              ( ['in'], c_int, 'Value' ),
              ( ['in'], c_int, 'Mask' )),
    COMMETHOD([dispid(839)], HRESULT, 'SetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Value' )),
]
################################################################
## code template for ITextFont2 implementation
##class ITextFont2_Impl(object):
##    def GetPropertyInfo(self, Index):
##        '-no docstring-'
##        #return pType, pValue
##
##    def GetProperty(self, Type):
##        '-no docstring-'
##        #return pValue
##
##    def IsEqual2(self, pFont):
##        '-no docstring-'
##        #return pB
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CompressionMode = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CharRep = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def LinkType(self):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    MathZone = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return ppFont
##    def _set(self, ppFont):
##        '-no docstring-'
##    Duplicate2 = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    SpaceExtension = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Overlapping = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ModWidthSpace = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Scaling = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Cookie = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    AutoLigatures = property(_get, _set, doc = _set.__doc__)
##
##    def GetEffects2(self):
##        '-no docstring-'
##        #return pValue, pMask
##
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pCount
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    UnderlinePositionMode = property(_get, _set, doc = _set.__doc__)
##
##    def SetEffects2(self, Value, Mask):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    AutospaceParens = property(_get, _set, doc = _set.__doc__)
##
##    def SetEffects(self, Value, Mask):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    AutospaceNumeric = property(_get, _set, doc = _set.__doc__)
##
##    def GetEffects(self):
##        '-no docstring-'
##        #return pValue, pMask
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    AutospaceAlpha = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ModWidthPairs = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    DoubleStrike = property(_get, _set, doc = _set.__doc__)
##
##    def SetProperty(self, Type, Value):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    OldNumbers = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    PositionSubSuper = property(_get, _set, doc = _set.__doc__)
##

class ITextDocument(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{8CC497C0-A1DF-11CE-8098-00AA0047BE5D}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextDocument._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Name',
              ( ['retval', 'out'], POINTER(BSTR), 'pName' )),
    COMMETHOD([dispid(1), 'propget'], HRESULT, 'Selection',
              ( ['retval', 'out'], POINTER(POINTER(ITextSelection)), 'ppSel' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'StoryCount',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(3), 'propget'], HRESULT, 'StoryRanges',
              ( ['retval', 'out'], POINTER(POINTER(ITextStoryRanges)), 'ppStories' )),
    COMMETHOD([dispid(4), 'propget'], HRESULT, 'Saved',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(4), 'propput'], HRESULT, 'Saved',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(5), 'propget'], HRESULT, 'DefaultTabStop',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(5), 'propput'], HRESULT, 'DefaultTabStop',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(6)], HRESULT, 'New'),
    COMMETHOD([dispid(7)], HRESULT, 'Open',
              ( ['in'], POINTER(VARIANT), 'pVar' ),
              ( ['in'], c_int, 'Flags' ),
              ( ['in'], c_int, 'CodePage' )),
    COMMETHOD([dispid(8)], HRESULT, 'Save',
              ( ['in'], POINTER(VARIANT), 'pVar' ),
              ( ['in'], c_int, 'Flags' ),
              ( ['in'], c_int, 'CodePage' )),
    COMMETHOD([dispid(9)], HRESULT, 'Freeze',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(10)], HRESULT, 'Unfreeze',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(11)], HRESULT, 'BeginEditCollection'),
    COMMETHOD([dispid(12)], HRESULT, 'EndEditCollection'),
    COMMETHOD([dispid(13)], HRESULT, 'Undo',
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(14)], HRESULT, 'Redo',
              ( ['in'], c_int, 'Count' ),
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(15)], HRESULT, 'Range',
              ( ['in'], c_int, 'cpActive' ),
              ( ['in'], c_int, 'cpAnchor' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange)), 'ppRange' )),
    COMMETHOD([dispid(16)], HRESULT, 'RangeFromPoint',
              ( ['in'], c_int, 'x' ),
              ( ['in'], c_int, 'y' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange)), 'ppRange' )),
]
################################################################
## code template for ITextDocument implementation
##class ITextDocument_Impl(object):
##    def Redo(self, Count):
##        '-no docstring-'
##        #return pCount
##
##    @property
##    def Selection(self):
##        '-no docstring-'
##        #return ppSel
##
##    def BeginEditCollection(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Name(self):
##        '-no docstring-'
##        #return pName
##
##    def EndEditCollection(self):
##        '-no docstring-'
##        #return 
##
##    def Open(self, pVar, Flags, CodePage):
##        '-no docstring-'
##        #return 
##
##    def Undo(self, Count):
##        '-no docstring-'
##        #return pCount
##
##    def Freeze(self):
##        '-no docstring-'
##        #return pCount
##
##    def Range(self, cpActive, cpAnchor):
##        '-no docstring-'
##        #return ppRange
##
##    def Unfreeze(self):
##        '-no docstring-'
##        #return pCount
##
##    def RangeFromPoint(self, x, y):
##        '-no docstring-'
##        #return ppRange
##
##    @property
##    def StoryCount(self):
##        '-no docstring-'
##        #return pCount
##
##    def New(self):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    DefaultTabStop = property(_get, _set, doc = _set.__doc__)
##
##    def Save(self, pVar, Flags, CodePage):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Saved = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def StoryRanges(self):
##        '-no docstring-'
##        #return ppStories
##

class ITextRow(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5EF-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextRow._methods_ = [
    COMMETHOD([dispid(2048), 'propget'], HRESULT, 'Alignment',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2048), 'propput'], HRESULT, 'Alignment',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2049), 'propget'], HRESULT, 'CellCount',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2049), 'propput'], HRESULT, 'CellCount',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2050), 'propget'], HRESULT, 'CellCountCache',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2050), 'propput'], HRESULT, 'CellCountCache',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2051), 'propget'], HRESULT, 'CellIndex',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2051), 'propput'], HRESULT, 'CellIndex',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2052), 'propget'], HRESULT, 'CellMargin',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2052), 'propput'], HRESULT, 'CellMargin',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2053), 'propget'], HRESULT, 'Height',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2053), 'propput'], HRESULT, 'Height',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2054), 'propget'], HRESULT, 'Indent',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2054), 'propput'], HRESULT, 'Indent',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2055), 'propget'], HRESULT, 'KeepTogether',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2055), 'propput'], HRESULT, 'KeepTogether',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2056), 'propget'], HRESULT, 'KeepWithNext',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2056), 'propput'], HRESULT, 'KeepWithNext',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2057), 'propget'], HRESULT, 'NestLevel',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2058), 'propget'], HRESULT, 'RTL',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2058), 'propput'], HRESULT, 'RTL',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2080), 'propget'], HRESULT, 'CellAlignment',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2080), 'propput'], HRESULT, 'CellAlignment',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2081), 'propget'], HRESULT, 'CellColorBack',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2081), 'propput'], HRESULT, 'CellColorBack',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2082), 'propget'], HRESULT, 'CellColorFore',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2082), 'propput'], HRESULT, 'CellColorFore',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2083), 'propget'], HRESULT, 'CellMergeFlags',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2083), 'propput'], HRESULT, 'CellMergeFlags',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2084), 'propget'], HRESULT, 'CellShading',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2084), 'propput'], HRESULT, 'CellShading',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2085), 'propget'], HRESULT, 'CellVerticalText',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2085), 'propput'], HRESULT, 'CellVerticalText',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2086), 'propget'], HRESULT, 'CellWidth',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2086), 'propput'], HRESULT, 'CellWidth',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(2096)], HRESULT, 'GetCellBorderColors',
              ( ['out'], POINTER(c_int), 'pcrLeft' ),
              ( ['out'], POINTER(c_int), 'pcrTop' ),
              ( ['out'], POINTER(c_int), 'pcrRight' ),
              ( ['out'], POINTER(c_int), 'pcrBottom' )),
    COMMETHOD([dispid(2097)], HRESULT, 'GetCellBorderWidths',
              ( ['out'], POINTER(c_int), 'pduLeft' ),
              ( ['out'], POINTER(c_int), 'pduTop' ),
              ( ['out'], POINTER(c_int), 'pduRight' ),
              ( ['out'], POINTER(c_int), 'pduBottom' )),
    COMMETHOD([dispid(2098)], HRESULT, 'SetCellBorderColors',
              ( ['in'], c_int, 'crLeft' ),
              ( ['in'], c_int, 'crTop' ),
              ( ['in'], c_int, 'crRight' ),
              ( ['in'], c_int, 'crBottom' )),
    COMMETHOD([dispid(2099)], HRESULT, 'SetCellBorderWidths',
              ( ['in'], c_int, 'duLeft' ),
              ( ['in'], c_int, 'duTop' ),
              ( ['in'], c_int, 'duRight' ),
              ( ['in'], c_int, 'duBottom' )),
    COMMETHOD([dispid(2112)], HRESULT, 'Apply',
              ( ['in'], c_int, 'cRow' ),
              ( ['in'], c_int, 'Flags' )),
    COMMETHOD([dispid(2113)], HRESULT, 'CanChange',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2114)], HRESULT, 'GetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(2115)], HRESULT, 'Insert',
              ( ['in'], c_int, 'cRow' )),
    COMMETHOD([dispid(2116)], HRESULT, 'IsEqual',
              ( ['in'], POINTER(ITextRow), 'pRow' ),
              ( ['retval', 'out'], POINTER(c_int), 'pB' )),
    COMMETHOD([dispid(2117)], HRESULT, 'Reset',
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(2118)], HRESULT, 'SetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Value' )),
]
################################################################
## code template for ITextRow implementation
##class ITextRow_Impl(object):
##    def GetProperty(self, Type):
##        '-no docstring-'
##        #return pValue
##
##    def Insert(self, cRow):
##        '-no docstring-'
##        #return 
##
##    def GetCellBorderWidths(self):
##        '-no docstring-'
##        #return pduLeft, pduTop, pduRight, pduBottom
##
##    def CanChange(self):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellColorFore = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellCountCache = property(_get, _set, doc = _set.__doc__)
##
##    def GetCellBorderColors(self):
##        '-no docstring-'
##        #return pcrLeft, pcrTop, pcrRight, pcrBottom
##
##    def IsEqual(self, pRow):
##        '-no docstring-'
##        #return pB
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellIndex = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    RTL = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def NestLevel(self):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellColorBack = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellMergeFlags = property(_get, _set, doc = _set.__doc__)
##
##    def Reset(self, Value):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellWidth = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Indent = property(_get, _set, doc = _set.__doc__)
##
##    def SetCellBorderColors(self, crLeft, crTop, crRight, crBottom):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellCount = property(_get, _set, doc = _set.__doc__)
##
##    def SetCellBorderWidths(self, duLeft, duTop, duRight, duBottom):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    KeepTogether = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellShading = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    KeepWithNext = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellAlignment = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Height = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellVerticalText = property(_get, _set, doc = _set.__doc__)
##
##    def SetProperty(self, Type, Value):
##        '-no docstring-'
##        #return 
##
##    def Apply(self, cRow, Flags):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    CellMargin = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Alignment = property(_get, _set, doc = _set.__doc__)
##

tagSTATSTG._fields_ = [
    ('pwcsName', WSTRING),
    ('Type', c_ulong),
    ('cbSize', _ULARGE_INTEGER),
    ('mtime', _FILETIME),
    ('ctime', _FILETIME),
    ('atime', _FILETIME),
    ('grfMode', c_ulong),
    ('grfLocksSupported', c_ulong),
    ('clsid', comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.GUID),
    ('grfStateBits', c_ulong),
    ('reserved', c_ulong),
]
assert sizeof(tagSTATSTG) == 72, sizeof(tagSTATSTG)
assert alignment(tagSTATSTG) == 8, alignment(tagSTATSTG)
ITextPara._methods_ = [
    COMMETHOD([dispid(0), 'propget'], HRESULT, 'Duplicate',
              ( ['retval', 'out'], POINTER(POINTER(ITextPara)), 'ppPara' )),
    COMMETHOD([dispid(0), 'propput'], HRESULT, 'Duplicate',
              ( ['in'], POINTER(ITextPara), 'ppPara' )),
    COMMETHOD([dispid(1025)], HRESULT, 'CanChange',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1026)], HRESULT, 'IsEqual',
              ( ['in'], POINTER(ITextPara), 'pPara' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1027)], HRESULT, 'Reset',
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(1028), 'propget'], HRESULT, 'Style',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1028), 'propput'], HRESULT, 'Style',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1029), 'propget'], HRESULT, 'Alignment',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1029), 'propput'], HRESULT, 'Alignment',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1030), 'propget'], HRESULT, 'Hyphenation',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1030), 'propput'], HRESULT, 'Hyphenation',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1031), 'propget'], HRESULT, 'FirstLineIndent',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1032), 'propget'], HRESULT, 'KeepTogether',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1032), 'propput'], HRESULT, 'KeepTogether',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1033), 'propget'], HRESULT, 'KeepWithNext',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1033), 'propput'], HRESULT, 'KeepWithNext',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1040), 'propget'], HRESULT, 'LeftIndent',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1041), 'propget'], HRESULT, 'LineSpacing',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1042), 'propget'], HRESULT, 'LineSpacingRule',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1043), 'propget'], HRESULT, 'ListAlignment',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1043), 'propput'], HRESULT, 'ListAlignment',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1044), 'propget'], HRESULT, 'ListLevelIndex',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1044), 'propput'], HRESULT, 'ListLevelIndex',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1045), 'propget'], HRESULT, 'ListStart',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1045), 'propput'], HRESULT, 'ListStart',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1046), 'propget'], HRESULT, 'ListTab',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1046), 'propput'], HRESULT, 'ListTab',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(1047), 'propget'], HRESULT, 'ListType',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1047), 'propput'], HRESULT, 'ListType',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1048), 'propget'], HRESULT, 'NoLineNumber',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1048), 'propput'], HRESULT, 'NoLineNumber',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1049), 'propget'], HRESULT, 'PageBreakBefore',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1049), 'propput'], HRESULT, 'PageBreakBefore',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1056), 'propget'], HRESULT, 'RightIndent',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1056), 'propput'], HRESULT, 'RightIndent',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(1057)], HRESULT, 'SetIndents',
              ( ['in'], c_float, 'First' ),
              ( ['in'], c_float, 'Left' ),
              ( ['in'], c_float, 'Right' )),
    COMMETHOD([dispid(1058)], HRESULT, 'SetLineSpacing',
              ( ['in'], c_int, 'Rule' ),
              ( ['in'], c_float, 'Spacing' )),
    COMMETHOD([dispid(1059), 'propget'], HRESULT, 'SpaceAfter',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1059), 'propput'], HRESULT, 'SpaceAfter',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(1060), 'propget'], HRESULT, 'SpaceBefore',
              ( ['retval', 'out'], POINTER(c_float), 'pValue' )),
    COMMETHOD([dispid(1060), 'propput'], HRESULT, 'SpaceBefore',
              ( ['in'], c_float, 'pValue' )),
    COMMETHOD([dispid(1061), 'propget'], HRESULT, 'WidowControl',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1061), 'propput'], HRESULT, 'WidowControl',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1062), 'propget'], HRESULT, 'TabCount',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(1063)], HRESULT, 'AddTab',
              ( ['in'], c_float, 'tbPos' ),
              ( ['in'], c_int, 'tbAlign' ),
              ( ['in'], c_int, 'tbLeader' )),
    COMMETHOD([dispid(1064)], HRESULT, 'ClearAllTabs'),
    COMMETHOD([dispid(1065)], HRESULT, 'DeleteTab',
              ( ['in'], c_float, 'tbPos' )),
    COMMETHOD([dispid(1072)], HRESULT, 'GetTab',
              ( ['in'], c_int, 'iTab' ),
              ( ['out'], POINTER(c_float), 'ptbPos' ),
              ( ['out'], POINTER(c_int), 'ptbAlign' ),
              ( ['out'], POINTER(c_int), 'ptbLeader' )),
]
################################################################
## code template for ITextPara implementation
##class ITextPara_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Style = property(_get, _set, doc = _set.__doc__)
##
##    def SetLineSpacing(self, Rule, Spacing):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ListStart = property(_get, _set, doc = _set.__doc__)
##
##    def CanChange(self):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return ppPara
##    def _set(self, ppPara):
##        '-no docstring-'
##    Duplicate = property(_get, _set, doc = _set.__doc__)
##
##    def IsEqual(self, pPara):
##        '-no docstring-'
##        #return pValue
##
##    def AddTab(self, tbPos, tbAlign, tbLeader):
##        '-no docstring-'
##        #return 
##
##    def ClearAllTabs(self):
##        '-no docstring-'
##        #return 
##
##    @property
##    def TabCount(self):
##        '-no docstring-'
##        #return pCount
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    WidowControl = property(_get, _set, doc = _set.__doc__)
##
##    def DeleteTab(self, tbPos):
##        '-no docstring-'
##        #return 
##
##    @property
##    def LineSpacing(self):
##        '-no docstring-'
##        #return pValue
##
##    @property
##    def FirstLineIndent(self):
##        '-no docstring-'
##        #return pValue
##
##    @property
##    def LeftIndent(self):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ListAlignment = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    RightIndent = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def LineSpacingRule(self):
##        '-no docstring-'
##        #return pValue
##
##    def Reset(self, Value):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    NoLineNumber = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    KeepTogether = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ListType = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    SpaceAfter = property(_get, _set, doc = _set.__doc__)
##
##    def SetIndents(self, First, Left, Right):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    SpaceBefore = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    KeepWithNext = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ListTab = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Hyphenation = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    ListLevelIndex = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    PageBreakBefore = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Alignment = property(_get, _set, doc = _set.__doc__)
##
##    def GetTab(self, iTab):
##        '-no docstring-'
##        #return ptbPos, ptbAlign, ptbLeader
##

class ITextSelection2(ITextRange2):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E1-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class ITextPara2(ITextPara):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextRange2._methods_ = [
    COMMETHOD([dispid(580), 'propget'], HRESULT, 'Cch',
              ( ['retval', 'out'], POINTER(c_int), 'pcch' )),
    COMMETHOD([dispid(581), 'propget'], HRESULT, 'Cells',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppCells' )),
    COMMETHOD([dispid(582), 'propget'], HRESULT, 'Column',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppColumn' )),
    COMMETHOD([dispid(583), 'propget'], HRESULT, 'Count',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(584), 'propget'], HRESULT, 'Duplicate2',
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
    COMMETHOD([dispid(585), 'propget'], HRESULT, 'Font2',
              ( ['retval', 'out'], POINTER(POINTER(ITextFont2)), 'ppFont' )),
    COMMETHOD([dispid(585), 'propput'], HRESULT, 'Font2',
              ( ['in'], POINTER(ITextFont2), 'ppFont' )),
    COMMETHOD([dispid(586), 'propget'], HRESULT, 'FormattedText2',
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
    COMMETHOD([dispid(586), 'propput'], HRESULT, 'FormattedText2',
              ( ['in'], POINTER(ITextRange2), 'ppRange' )),
    COMMETHOD([dispid(587), 'propget'], HRESULT, 'Gravity',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(587), 'propput'], HRESULT, 'Gravity',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(588), 'propget'], HRESULT, 'Para2',
              ( ['retval', 'out'], POINTER(POINTER(ITextPara2)), 'ppPara' )),
    COMMETHOD([dispid(588), 'propput'], HRESULT, 'Para2',
              ( ['in'], POINTER(ITextPara2), 'ppPara' )),
    COMMETHOD([dispid(589), 'propget'], HRESULT, 'Row',
              ( ['retval', 'out'], POINTER(POINTER(ITextRow)), 'ppRow' )),
    COMMETHOD([dispid(590), 'propget'], HRESULT, 'StartPara',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(591), 'propget'], HRESULT, 'Table',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppTable' )),
    COMMETHOD([dispid(592), 'propget'], HRESULT, 'URL',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(592), 'propput'], HRESULT, 'URL',
              ( ['in'], BSTR, 'pbstr' )),
    COMMETHOD([dispid(608)], HRESULT, 'AddSubrange',
              ( ['in'], c_int, 'cp1' ),
              ( ['in'], c_int, 'cp2' ),
              ( ['in'], c_int, 'Activate' )),
    COMMETHOD([dispid(609)], HRESULT, 'BuildUpMath',
              ( ['in'], c_int, 'Flags' )),
    COMMETHOD([dispid(610)], HRESULT, 'DeleteSubrange',
              ( ['in'], c_int, 'cpFirst' ),
              ( ['in'], c_int, 'cpLim' )),
    COMMETHOD([dispid(611)], HRESULT, 'Find',
              ( ['in'], POINTER(ITextRange2), 'pRange' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'Flags' ),
              ( ['out'], POINTER(c_int), 'pDelta' )),
    COMMETHOD([dispid(612)], HRESULT, 'GetChar2',
              ( ['out'], POINTER(c_int), 'pChar' ),
              ( ['in'], c_int, 'Offset' )),
    COMMETHOD([dispid(613)], HRESULT, 'GetDropCap',
              ( ['out'], POINTER(c_int), 'pcLine' ),
              ( ['out'], POINTER(c_int), 'pPosition' )),
    COMMETHOD([dispid(614)], HRESULT, 'GetInlineObject',
              ( ['out'], POINTER(c_int), 'pType' ),
              ( ['out'], POINTER(c_int), 'pAlign' ),
              ( ['out'], POINTER(c_int), 'pChar' ),
              ( ['out'], POINTER(c_int), 'pChar1' ),
              ( ['out'], POINTER(c_int), 'pChar2' ),
              ( ['out'], POINTER(c_int), 'pCount' ),
              ( ['out'], POINTER(c_int), 'pTeXStyle' ),
              ( ['out'], POINTER(c_int), 'pcCol' ),
              ( ['out'], POINTER(c_int), 'pLevel' )),
    COMMETHOD([dispid(615)], HRESULT, 'GetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(616)], HRESULT, 'GetRect',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'pLeft' ),
              ( ['out'], POINTER(c_int), 'pTop' ),
              ( ['out'], POINTER(c_int), 'pRight' ),
              ( ['out'], POINTER(c_int), 'pBottom' ),
              ( ['out'], POINTER(c_int), 'pHit' )),
    COMMETHOD([dispid(617)], HRESULT, 'GetSubrange',
              ( ['in'], c_int, 'iSubrange' ),
              ( ['out'], POINTER(c_int), 'pcpFirst' ),
              ( ['out'], POINTER(c_int), 'pcpLim' )),
    COMMETHOD([dispid(618)], HRESULT, 'GetText2',
              ( ['in'], c_int, 'Flags' ),
              ( ['out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(619)], HRESULT, 'HexToUnicode'),
    COMMETHOD([dispid(620)], HRESULT, 'InsertTable',
              ( ['in'], c_int, 'cCol' ),
              ( ['in'], c_int, 'cRow' ),
              ( ['in'], c_int, 'AutoFit' )),
    COMMETHOD([dispid(621)], HRESULT, 'Linearize',
              ( ['in'], c_int, 'Flags' )),
    COMMETHOD([dispid(622)], HRESULT, 'SetActiveSubrange',
              ( ['in'], c_int, 'cpAnchor' ),
              ( ['in'], c_int, 'cpActive' )),
    COMMETHOD([dispid(623)], HRESULT, 'SetDropCap',
              ( ['in'], c_int, 'cLine' ),
              ( ['in'], c_int, 'Position' )),
    COMMETHOD([dispid(624)], HRESULT, 'SetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(625)], HRESULT, 'SetText2',
              ( ['in'], c_int, 'Flags' ),
              ( ['in'], BSTR, 'bstr' )),
    COMMETHOD([dispid(626)], HRESULT, 'UnicodeToHex'),
    COMMETHOD([dispid(627)], HRESULT, 'SetInlineObject',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Align' ),
              ( ['in'], c_int, 'Char' ),
              ( ['in'], c_int, 'Char1' ),
              ( ['in'], c_int, 'Char2' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'TeXStyle' ),
              ( ['in'], c_int, 'cCol' )),
    COMMETHOD([dispid(628)], HRESULT, 'GetMathFunctionType',
              ( ['in'], BSTR, 'bstr' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(629)], HRESULT, 'InsertImage',
              ( ['in'], c_int, 'width' ),
              ( ['in'], c_int, 'Height' ),
              ( ['in'], c_int, 'ascent' ),
              ( ['in'], c_int, 'Type' ),
              ( ['in'], BSTR, 'bstrAltText' ),
              ( ['in'], POINTER(IStream), 'pStream' )),
]
################################################################
## code template for ITextRange2 implementation
##class ITextRange2_Impl(object):
##    def GetText2(self, Flags):
##        '-no docstring-'
##        #return pbstr
##
##    def GetProperty(self, Type):
##        '-no docstring-'
##        #return pValue
##
##    def InsertImage(self, width, Height, ascent, Type, bstrAltText, pStream):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return ppFont
##    def _set(self, ppFont):
##        '-no docstring-'
##    Font2 = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Gravity = property(_get, _set, doc = _set.__doc__)
##
##    def InsertTable(self, cCol, cRow, AutoFit):
##        '-no docstring-'
##        #return 
##
##    def DeleteSubrange(self, cpFirst, cpLim):
##        '-no docstring-'
##        #return 
##
##    def HexToUnicode(self):
##        '-no docstring-'
##        #return 
##
##    def SetDropCap(self, cLine, Position):
##        '-no docstring-'
##        #return 
##
##    def GetInlineObject(self):
##        '-no docstring-'
##        #return pType, pAlign, pChar, pChar1, pChar2, pCount, pTeXStyle, pcCol, pLevel
##
##    def _get(self):
##        '-no docstring-'
##        #return ppRange
##    def _set(self, ppRange):
##        '-no docstring-'
##    FormattedText2 = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Duplicate2(self):
##        '-no docstring-'
##        #return ppRange
##
##    def GetSubrange(self, iSubrange):
##        '-no docstring-'
##        #return pcpFirst, pcpLim
##
##    def AddSubrange(self, cp1, cp2, Activate):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pbstr
##    def _set(self, pbstr):
##        '-no docstring-'
##    URL = property(_get, _set, doc = _set.__doc__)
##
##    def BuildUpMath(self, Flags):
##        '-no docstring-'
##        #return 
##
##    def GetMathFunctionType(self, bstr):
##        '-no docstring-'
##        #return pValue
##
##    @property
##    def StartPara(self):
##        '-no docstring-'
##        #return pValue
##
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pCount
##
##    def Linearize(self, Flags):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Column(self):
##        '-no docstring-'
##        #return ppColumn
##
##    def GetRect(self, Type):
##        '-no docstring-'
##        #return pLeft, pTop, pRight, pBottom, pHit
##
##    @property
##    def Cells(self):
##        '-no docstring-'
##        #return ppCells
##
##    def _get(self):
##        '-no docstring-'
##        #return ppPara
##    def _set(self, ppPara):
##        '-no docstring-'
##    Para2 = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Cch(self):
##        '-no docstring-'
##        #return pcch
##
##    def SetInlineObject(self, Type, Align, Char, Char1, Char2, Count, TeXStyle, cCol):
##        '-no docstring-'
##        #return 
##
##    def SetText2(self, Flags, bstr):
##        '-no docstring-'
##        #return 
##
##    def GetDropCap(self):
##        '-no docstring-'
##        #return pcLine, pPosition
##
##    def SetActiveSubrange(self, cpAnchor, cpActive):
##        '-no docstring-'
##        #return 
##
##    def UnicodeToHex(self):
##        '-no docstring-'
##        #return 
##
##    def GetChar2(self, Offset):
##        '-no docstring-'
##        #return pChar
##
##    def SetProperty(self, Type, Value):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Table(self):
##        '-no docstring-'
##        #return ppTable
##
##    def Find(self, pRange, Count, Flags):
##        '-no docstring-'
##        #return pDelta
##
##    @property
##    def Row(self):
##        '-no docstring-'
##        #return ppRow
##

ITextSelection2._methods_ = [
]
################################################################
## code template for ITextSelection2 implementation
##class ITextSelection2_Impl(object):

class ITextStrings(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IDispatch):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E7-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
ITextStrings._methods_ = [
    COMMETHOD([dispid(0)], HRESULT, 'Item',
              ( ['in'], c_int, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
    COMMETHOD([dispid(2), 'propget'], HRESULT, 'Count',
              ( ['retval', 'out'], POINTER(c_int), 'pCount' )),
    COMMETHOD([dispid(3)], HRESULT, 'Add',
              ( ['in'], BSTR, 'bstr' )),
    COMMETHOD([dispid(4)], HRESULT, 'Append',
              ( ['in'], POINTER(ITextRange2), 'pRange' ),
              ( ['in'], c_int, 'iString' )),
    COMMETHOD([dispid(5)], HRESULT, 'Cat2',
              ( ['in'], c_int, 'iString' )),
    COMMETHOD([dispid(6)], HRESULT, 'CatTop2',
              ( ['in'], BSTR, 'bstr' )),
    COMMETHOD([dispid(7)], HRESULT, 'DeleteRange',
              ( ['in'], POINTER(ITextRange2), 'pRange' )),
    COMMETHOD([dispid(8)], HRESULT, 'EncodeFunction',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Align' ),
              ( ['in'], c_int, 'Char' ),
              ( ['in'], c_int, 'Char1' ),
              ( ['in'], c_int, 'Char2' ),
              ( ['in'], c_int, 'Count' ),
              ( ['in'], c_int, 'TeXStyle' ),
              ( ['in'], c_int, 'cCol' ),
              ( ['in'], POINTER(ITextRange2), 'pRange' )),
    COMMETHOD([dispid(9)], HRESULT, 'GetCch',
              ( ['in'], c_int, 'iString' ),
              ( ['out'], POINTER(c_int), 'pcch' )),
    COMMETHOD([dispid(10)], HRESULT, 'InsertNullStr',
              ( ['in'], c_int, 'iString' )),
    COMMETHOD([dispid(11)], HRESULT, 'MoveBoundary',
              ( ['in'], c_int, 'iString' ),
              ( ['in'], c_int, 'Cch' )),
    COMMETHOD([dispid(12)], HRESULT, 'PrefixTop',
              ( ['in'], BSTR, 'bstr' )),
    COMMETHOD([dispid(13)], HRESULT, 'Remove',
              ( ['in'], c_int, 'iString' ),
              ( ['in'], c_int, 'cString' )),
    COMMETHOD([dispid(14)], HRESULT, 'SetFormattedText',
              ( ['in'], POINTER(ITextRange2), 'pRangeD' ),
              ( ['in'], POINTER(ITextRange2), 'pRangeS' )),
    COMMETHOD([dispid(15)], HRESULT, 'SetOpCp',
              ( ['in'], c_int, 'iString' ),
              ( ['in'], c_int, 'cp' )),
    COMMETHOD([dispid(16)], HRESULT, 'SuffixTop',
              ( ['in'], BSTR, 'bstr' ),
              ( ['in'], POINTER(ITextRange2), 'pRange' )),
    COMMETHOD([dispid(17)], HRESULT, 'Swap'),
]
################################################################
## code template for ITextStrings implementation
##class ITextStrings_Impl(object):
##    @property
##    def Count(self):
##        '-no docstring-'
##        #return pCount
##
##    def SetOpCp(self, iString, cp):
##        '-no docstring-'
##        #return 
##
##    def EncodeFunction(self, Type, Align, Char, Char1, Char2, Count, TeXStyle, cCol, pRange):
##        '-no docstring-'
##        #return 
##
##    def SuffixTop(self, bstr, pRange):
##        '-no docstring-'
##        #return 
##
##    def MoveBoundary(self, iString, Cch):
##        '-no docstring-'
##        #return 
##
##    def DeleteRange(self, pRange):
##        '-no docstring-'
##        #return 
##
##    def InsertNullStr(self, iString):
##        '-no docstring-'
##        #return 
##
##    def CatTop2(self, bstr):
##        '-no docstring-'
##        #return 
##
##    def Remove(self, iString, cString):
##        '-no docstring-'
##        #return 
##
##    def Cat2(self, iString):
##        '-no docstring-'
##        #return 
##
##    def Item(self, Index):
##        '-no docstring-'
##        #return ppRange
##
##    def Add(self, bstr):
##        '-no docstring-'
##        #return 
##
##    def Swap(self):
##        '-no docstring-'
##        #return 
##
##    def SetFormattedText(self, pRangeD, pRangeS):
##        '-no docstring-'
##        #return 
##
##    def GetCch(self, iString):
##        '-no docstring-'
##        #return pcch
##
##    def PrefixTop(self, bstr):
##        '-no docstring-'
##        #return 
##
##    def Append(self, pRange, iString):
##        '-no docstring-'
##        #return 
##

class ITextDocument2(ITextDocument):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5E0-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['dual', 'nonextensible', 'oleautomation']
class ITextStory(comtypes.gen._00020430_0000_0000_C000_000000000046_0_2_0.IUnknown):
    _case_insensitive_ = True
    _iid_ = GUID('{C241F5F3-7206-11D8-A2C7-00A0D1D6C6B3}')
    _idlflags_ = ['nonextensible']
ITextDocument2._methods_ = [
    COMMETHOD([dispid(17), helpstring(u'method GetCaretType'), 'propget'], HRESULT, 'CaretType',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(17), helpstring(u'method GetCaretType'), 'propput'], HRESULT, 'CaretType',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(18), helpstring(u'method GetDisplays'), 'propget'], HRESULT, 'Displays',
              ( ['retval', 'out'], POINTER(POINTER(ITextDisplays)), 'ppDisplays' )),
    COMMETHOD([dispid(19), helpstring(u'method GetDocumentFont'), 'propget'], HRESULT, 'DocumentFont',
              ( ['retval', 'out'], POINTER(POINTER(ITextFont2)), 'ppFont' )),
    COMMETHOD([dispid(19), helpstring(u'method GetDocumentFont'), 'propput'], HRESULT, 'DocumentFont',
              ( ['in'], POINTER(ITextFont2), 'ppFont' )),
    COMMETHOD([dispid(20), helpstring(u'method GetDocumentPara'), 'propget'], HRESULT, 'DocumentPara',
              ( ['retval', 'out'], POINTER(POINTER(ITextPara2)), 'ppPara' )),
    COMMETHOD([dispid(20), helpstring(u'method GetDocumentPara'), 'propput'], HRESULT, 'DocumentPara',
              ( ['in'], POINTER(ITextPara2), 'ppPara' )),
    COMMETHOD([dispid(21), helpstring(u'method GetEastAsianFlags'), 'propget'], HRESULT, 'EastAsianFlags',
              ( ['retval', 'out'], POINTER(c_int), 'pFlags' )),
    COMMETHOD([dispid(22), helpstring(u'method GetGenerator'), 'propget'], HRESULT, 'Generator',
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([dispid(23), helpstring(u'method SetIMEInProgress'), 'propput'], HRESULT, 'IMEInProgress',
              ( ['in'], c_int, 'rhs' )),
    COMMETHOD([dispid(24), helpstring(u'method GetNotificationMode'), 'propget'], HRESULT, 'NotificationMode',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(24), helpstring(u'method GetNotificationMode'), 'propput'], HRESULT, 'NotificationMode',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(25), helpstring(u'method GetSelection2'), 'propget'], HRESULT, 'Selection2',
              ( ['retval', 'out'], POINTER(POINTER(ITextSelection2)), 'ppSel' )),
    COMMETHOD([dispid(26), helpstring(u'method Selection2'), 'propget'], HRESULT, 'StoryRanges2',
              ( ['retval', 'out'], POINTER(POINTER(ITextStoryRanges2)), 'ppStories' )),
    COMMETHOD([dispid(27), helpstring(u'method GetTypographyOptions'), 'propget'], HRESULT, 'TypographyOptions',
              ( ['retval', 'out'], POINTER(c_int), 'pOptions' )),
    COMMETHOD([dispid(28), helpstring(u'method Version'), 'propget'], HRESULT, 'Version',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(29), helpstring(u'method GetWindow'), 'propget'], HRESULT, 'Window',
              ( ['retval', 'out'], POINTER(c_longlong), 'pHwnd' )),
    COMMETHOD([dispid(30), helpstring(u'method AttachMsgFilter')], HRESULT, 'AttachMsgFilter',
              ( ['in'], POINTER(IUnknown), 'pFilter' )),
    COMMETHOD([dispid(31), helpstring(u'method CheckTextLimit')], HRESULT, 'CheckTextLimit',
              ( [], c_int, 'Cch' ),
              ( [], POINTER(c_int), 'pcch' )),
    COMMETHOD([dispid(32), helpstring(u'method GetCallManager')], HRESULT, 'GetCallManager',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppVoid' )),
    COMMETHOD([dispid(33), helpstring(u'method GetClientRect')], HRESULT, 'GetClientRect',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'pLeft' ),
              ( ['out'], POINTER(c_int), 'pTop' ),
              ( ['out'], POINTER(c_int), 'pRight' ),
              ( ['out'], POINTER(c_int), 'pBottom' )),
    COMMETHOD([dispid(34), helpstring(u'method GetEffectColor')], HRESULT, 'GetEffectColor',
              ( ['in'], c_int, 'Index' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(35), helpstring(u'method GetImmContext')], HRESULT, 'GetImmContext',
              ( ['retval', 'out'], POINTER(c_longlong), 'pContext' )),
    COMMETHOD([dispid(36), helpstring(u'method GetPreferredFont')], HRESULT, 'GetPreferredFont',
              ( ['in'], c_int, 'cp' ),
              ( ['in'], c_int, 'CharRep' ),
              ( ['in'], c_int, 'Options' ),
              ( ['in'], c_int, 'curCharRep' ),
              ( ['in'], c_int, 'curFontSize' ),
              ( ['out'], POINTER(BSTR), 'pbstr' ),
              ( ['out'], POINTER(c_int), 'pPitchAndFamily' ),
              ( ['out'], POINTER(c_int), 'pNewFontSize' )),
    COMMETHOD([dispid(37), helpstring(u'method GetProperty')], HRESULT, 'GetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(38), helpstring(u'method GetStrings')], HRESULT, 'GetStrings',
              ( ['out'], POINTER(POINTER(ITextStrings)), 'ppStrs' )),
    COMMETHOD([dispid(39), helpstring(u'method Notify')], HRESULT, 'Notify',
              ( ['in'], c_int, 'Notify' )),
    COMMETHOD([dispid(40), helpstring(u'method Selection2')], HRESULT, 'Range2',
              ( ['in'], c_int, 'cpActive' ),
              ( ['in'], c_int, 'cpAnchor' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
    COMMETHOD([dispid(41), helpstring(u'method RangeFromPoint2')], HRESULT, 'RangeFromPoint2',
              ( ['in'], c_int, 'x' ),
              ( ['in'], c_int, 'y' ),
              ( ['in'], c_int, 'Type' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
    COMMETHOD([dispid(42), helpstring(u'method ReleaseCallManager')], HRESULT, 'ReleaseCallManager',
              ( ['in'], POINTER(IUnknown), 'pVoid' )),
    COMMETHOD([dispid(43), helpstring(u'method ReleaseImmContext')], HRESULT, 'ReleaseImmContext',
              ( ['in'], c_longlong, 'Context' )),
    COMMETHOD([dispid(44), helpstring(u'method SetEffectColor')], HRESULT, 'SetEffectColor',
              ( ['in'], c_int, 'Index' ),
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(45), helpstring(u'method SetProperty')], HRESULT, 'SetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(46), helpstring(u'method SetTypographyOptions')], HRESULT, 'SetTypographyOptions',
              ( ['in'], c_int, 'Options' ),
              ( ['in'], c_int, 'Mask' )),
    COMMETHOD([dispid(47), helpstring(u'method SysBeep')], HRESULT, 'SysBeep'),
    COMMETHOD([dispid(48), helpstring(u'method Update')], HRESULT, 'Update',
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([dispid(49), helpstring(u'method UpdateWindow')], HRESULT, 'UpdateWindow'),
    COMMETHOD([dispid(50), helpstring(u'method GetMathProperties')], HRESULT, 'GetMathProperties',
              ( ['out'], POINTER(c_int), 'pOptions' )),
    COMMETHOD([dispid(51), helpstring(u'method SetMathProperties')], HRESULT, 'SetMathProperties',
              ( ['in'], c_int, 'Options' ),
              ( ['in'], c_int, 'Mask' )),
    COMMETHOD([dispid(60), 'propget'], HRESULT, 'ActiveStory',
              ( ['retval', 'out'], POINTER(POINTER(ITextStory)), 'ppStory' )),
    COMMETHOD([dispid(60), 'propput'], HRESULT, 'ActiveStory',
              ( ['in'], POINTER(ITextStory), 'ppStory' )),
    COMMETHOD([dispid(61), 'propget'], HRESULT, 'MainStory',
              ( ['retval', 'out'], POINTER(POINTER(ITextStory)), 'ppStory' )),
    COMMETHOD([dispid(62), 'propget'], HRESULT, 'NewStory',
              ( ['retval', 'out'], POINTER(POINTER(ITextStory)), 'ppStory' )),
    COMMETHOD([dispid(66)], HRESULT, 'GetStory',
              ( ['in'], c_int, 'Index' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextStory)), 'ppStory' )),
]
################################################################
## code template for ITextDocument2 implementation
##class ITextDocument2_Impl(object):
##    def GetProperty(self, Type):
##        u'method GetProperty'
##        #return pValue
##
##    @property
##    def Selection2(self):
##        u'method GetSelection2'
##        #return ppSel
##
##    def UpdateWindow(self):
##        u'method UpdateWindow'
##        #return 
##
##    def RangeFromPoint2(self, x, y, Type):
##        u'method RangeFromPoint2'
##        #return ppRange
##
##    @property
##    def StoryRanges2(self):
##        u'method Selection2'
##        #return ppStories
##
##    @property
##    def Window(self):
##        u'method GetWindow'
##        #return pHwnd
##
##    @property
##    def Version(self):
##        u'method Version'
##        #return pValue
##
##    def GetPreferredFont(self, cp, CharRep, Options, curCharRep, curFontSize):
##        u'method GetPreferredFont'
##        #return pbstr, pPitchAndFamily, pNewFontSize
##
##    def _get(self):
##        u'method GetDocumentPara'
##        #return ppPara
##    def _set(self, ppPara):
##        u'method GetDocumentPara'
##    DocumentPara = property(_get, _set, doc = _set.__doc__)
##
##    def SysBeep(self):
##        u'method SysBeep'
##        #return 
##
##    def CheckTextLimit(self, Cch, pcch):
##        u'method CheckTextLimit'
##        #return 
##
##    def SetEffectColor(self, Index, Value):
##        u'method SetEffectColor'
##        #return 
##
##    @property
##    def Generator(self):
##        u'method GetGenerator'
##        #return pbstr
##
##    def GetEffectColor(self, Index):
##        u'method GetEffectColor'
##        #return pValue
##
##    def GetMathProperties(self):
##        u'method GetMathProperties'
##        #return pOptions
##
##    def Update(self, Value):
##        u'method Update'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return ppStory
##    def _set(self, ppStory):
##        '-no docstring-'
##    ActiveStory = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def EastAsianFlags(self):
##        u'method GetEastAsianFlags'
##        #return pFlags
##
##    def GetStrings(self):
##        u'method GetStrings'
##        #return ppStrs
##
##    def _set(self, rhs):
##        u'method SetIMEInProgress'
##    IMEInProgress = property(fset = _set, doc = _set.__doc__)
##
##    @property
##    def MainStory(self):
##        '-no docstring-'
##        #return ppStory
##
##    def SetMathProperties(self, Options, Mask):
##        u'method SetMathProperties'
##        #return 
##
##    def _get(self):
##        u'method GetCaretType'
##        #return pValue
##    def _set(self, pValue):
##        u'method GetCaretType'
##    CaretType = property(_get, _set, doc = _set.__doc__)
##
##    def GetStory(self, Index):
##        '-no docstring-'
##        #return ppStory
##
##    def Notify(self, Notify):
##        u'method Notify'
##        #return 
##
##    def ReleaseImmContext(self, Context):
##        u'method ReleaseImmContext'
##        #return 
##
##    def GetCallManager(self):
##        u'method GetCallManager'
##        #return ppVoid
##
##    @property
##    def NewStory(self):
##        '-no docstring-'
##        #return ppStory
##
##    def GetClientRect(self, Type):
##        u'method GetClientRect'
##        #return pLeft, pTop, pRight, pBottom
##
##    def SetTypographyOptions(self, Options, Mask):
##        u'method SetTypographyOptions'
##        #return 
##
##    def AttachMsgFilter(self, pFilter):
##        u'method AttachMsgFilter'
##        #return 
##
##    def GetImmContext(self):
##        u'method GetImmContext'
##        #return pContext
##
##    def ReleaseCallManager(self, pVoid):
##        u'method ReleaseCallManager'
##        #return 
##
##    def _get(self):
##        u'method GetNotificationMode'
##        #return pValue
##    def _set(self, pValue):
##        u'method GetNotificationMode'
##    NotificationMode = property(_get, _set, doc = _set.__doc__)
##
##    def _get(self):
##        u'method GetDocumentFont'
##        #return ppFont
##    def _set(self, ppFont):
##        u'method GetDocumentFont'
##    DocumentFont = property(_get, _set, doc = _set.__doc__)
##
##    def Range2(self, cpActive, cpAnchor):
##        u'method Selection2'
##        #return ppRange
##
##    @property
##    def Displays(self):
##        u'method GetDisplays'
##        #return ppDisplays
##
##    @property
##    def TypographyOptions(self):
##        u'method GetTypographyOptions'
##        #return pOptions
##
##    def SetProperty(self, Type, Value):
##        u'method SetProperty'
##        #return 
##

ITextStory._methods_ = [
    COMMETHOD(['propget'], HRESULT, 'Active',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD(['propput'], HRESULT, 'Active',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD(['propget'], HRESULT, 'Display',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppDisplay' )),
    COMMETHOD(['propget'], HRESULT, 'Index',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD(['propget'], HRESULT, 'Type',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD(['propput'], HRESULT, 'Type',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([], HRESULT, 'GetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([], HRESULT, 'GetRange',
              ( ['in'], c_int, 'cpActive' ),
              ( ['in'], c_int, 'cpAnchor' ),
              ( ['retval', 'out'], POINTER(POINTER(ITextRange2)), 'ppRange' )),
    COMMETHOD([], HRESULT, 'GetText',
              ( ['in'], c_int, 'Flags' ),
              ( ['retval', 'out'], POINTER(BSTR), 'pbstr' )),
    COMMETHOD([], HRESULT, 'SetFormattedText',
              ( ['in'], POINTER(IUnknown), 'pUnk' )),
    COMMETHOD([], HRESULT, 'SetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Value' )),
    COMMETHOD([], HRESULT, 'SetText',
              ( ['in'], c_int, 'Flags' ),
              ( ['in'], BSTR, 'bstr' )),
]
################################################################
## code template for ITextStory implementation
##class ITextStory_Impl(object):
##    @property
##    def Index(self):
##        '-no docstring-'
##        #return pValue
##
##    def GetRange(self, cpActive, cpAnchor):
##        '-no docstring-'
##        #return ppRange
##
##    def GetProperty(self, Type):
##        '-no docstring-'
##        #return pValue
##
##    def SetText(self, Flags, bstr):
##        '-no docstring-'
##        #return 
##
##    def GetText(self, Flags):
##        '-no docstring-'
##        #return pbstr
##
##    def SetProperty(self, Type, Value):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Active = property(_get, _set, doc = _set.__doc__)
##
##    def SetFormattedText(self, pUnk):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    Type = property(_get, _set, doc = _set.__doc__)
##
##    @property
##    def Display(self):
##        '-no docstring-'
##        #return ppDisplay
##

ITextPara2._methods_ = [
    COMMETHOD([dispid(1073), 'propget'], HRESULT, 'Borders',
              ( ['retval', 'out'], POINTER(POINTER(IUnknown)), 'ppBorders' )),
    COMMETHOD([dispid(1074), 'propget'], HRESULT, 'Duplicate2',
              ( ['retval', 'out'], POINTER(POINTER(ITextPara2)), 'ppPara' )),
    COMMETHOD([dispid(1074), 'propput'], HRESULT, 'Duplicate2',
              ( ['in'], POINTER(ITextPara2), 'ppPara' )),
    COMMETHOD([dispid(1075), 'propget'], HRESULT, 'FontAlignment',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1075), 'propput'], HRESULT, 'FontAlignment',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1076), 'propget'], HRESULT, 'HangingPunctuation',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1076), 'propput'], HRESULT, 'HangingPunctuation',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1077), 'propget'], HRESULT, 'SnapToGrid',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1077), 'propput'], HRESULT, 'SnapToGrid',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1078), 'propget'], HRESULT, 'TrimPunctuationAtStart',
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1078), 'propput'], HRESULT, 'TrimPunctuationAtStart',
              ( ['in'], c_int, 'pValue' )),
    COMMETHOD([dispid(1088)], HRESULT, 'GetEffects',
              ( ['out'], POINTER(c_int), 'pValue' ),
              ( ['out'], POINTER(c_int), 'pMask' )),
    COMMETHOD([dispid(1089)], HRESULT, 'GetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['retval', 'out'], POINTER(c_int), 'pValue' )),
    COMMETHOD([dispid(1090)], HRESULT, 'IsEqual2',
              ( ['in'], POINTER(ITextPara2), 'pPara' ),
              ( ['retval', 'out'], POINTER(c_int), 'pB' )),
    COMMETHOD([dispid(1091)], HRESULT, 'SetEffects',
              ( ['in'], c_int, 'Value' ),
              ( ['in'], c_int, 'Mask' )),
    COMMETHOD([dispid(1092)], HRESULT, 'SetProperty',
              ( ['in'], c_int, 'Type' ),
              ( ['in'], c_int, 'Value' )),
]
################################################################
## code template for ITextPara2 implementation
##class ITextPara2_Impl(object):
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    HangingPunctuation = property(_get, _set, doc = _set.__doc__)
##
##    def GetProperty(self, Type):
##        '-no docstring-'
##        #return pValue
##
##    def _get(self):
##        '-no docstring-'
##        #return ppPara
##    def _set(self, ppPara):
##        '-no docstring-'
##    Duplicate2 = property(_get, _set, doc = _set.__doc__)
##
##    def IsEqual2(self, pPara):
##        '-no docstring-'
##        #return pB
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    TrimPunctuationAtStart = property(_get, _set, doc = _set.__doc__)
##
##    def SetProperty(self, Type, Value):
##        '-no docstring-'
##        #return 
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    SnapToGrid = property(_get, _set, doc = _set.__doc__)
##
##    def SetEffects(self, Value, Mask):
##        '-no docstring-'
##        #return 
##
##    @property
##    def Borders(self):
##        '-no docstring-'
##        #return ppBorders
##
##    def GetEffects(self):
##        '-no docstring-'
##        #return pValue, pMask
##
##    def _get(self):
##        '-no docstring-'
##        #return pValue
##    def _set(self, pValue):
##        '-no docstring-'
##    FontAlignment = property(_get, _set, doc = _set.__doc__)
##

__all__ = [ 'tomLineSpaceMultiple', 'tomAllowFinalEOP',
           'tomEllipsisWord', 'tomListNumberAsSequence', 'tomHidden',
           'tomSpaceRelational', 'tomMathZoneOrdinary',
           'tomTruncateExisting', 'tomBracketsWithSeps',
           'tomUpperLimitAsSuperScript', 'ITextPara2',
           'tomPrimaryFooterStory', 'tomDontGrowWithContent',
           'tomGujarati', 'tomShiftJIS', 'tomIncludeInset',
           'tomShowDegPlaceHldr', 'tomMyanmar',
           'tomShowLLimPlaceHldr', 'tomFontPropTeXStyle',
           'tomMathInsColBefore', 'tomParaEffectDoNotHyphen',
           'tomMathParaAlignCenter', 'tomListNumberedHebrew',
           'tomWipeRight', 'tomSpaceUnary', 'tomFriendlyLinkName',
           'tomLeftSubSup', 'tomShowEmptyArgPlaceholders',
           'tomSelectionBlock', 'tomDocAutoLink', 'tomVTopCell',
           'tomMathBuildDown', 'tomMathDeleteArg2',
           'tomMathBrkBinSubPM', 'tomPhantomSmash',
           'tomMathDeleteArg1', 'tomLimitsSubSup',
           'tomMathDocEmptyArgMask', 'tomListNumberedThaiAlpha',
           'tomDashDotDot', 'tomMathRMargin', 'tomSpaceMask',
           'tomCanUndo', 'tomPC437', 'tomStack', 'tomAdjustCRLF',
           'MGREEK', 'tomPhantomASmash', 'tomApplyNow', 'tomLongDash',
           'tomTrackParms', 'tomLineSpaceAtLeast', 'tomBoxStrikeBLTR',
           'tomMathPostSpace', 'tomPhantomTransparent', 'tomMac',
           'tomHstring', 'tomCompressPunctuation', 'tomGetUtf16',
           'tomNoVpScroll', 'tomLisu', 'tomCreateNew',
           'tomCellStructureChangeOnly', 'tomPlain',
           'tomLinkProtected', 'tomWarichu', 'tomHeavyWave',
           'tomAlignDefault', 'tomUseOperandPrec',
           'tomBackgroundImage', 'tomListNumberedChT', 'tomObjectMax',
           'tomListNumberedChS', 'tomFontWeightThin', 'tomWave',
           'tomRuby', 'tomAlignCenter', 'tomMathArgShadingEnd',
           'tomCompressNone', 'tomDecDecSize', 'tomMathRichEdit',
           'tomMathBreakLeft', 'tomFontWeightMedium', 'tomImprint',
           'tomRoundedBoxHideBorder', 'tomEmboss', 'tomSection',
           'tomMathShiftTab', 'tomStretchBaseBelow',
           'tomUnderlinePositionBelow', 'tomMatrixAlignTopRow',
           'tomSelectionIP', 'tomFontStyleOblique', 'tomConvertRuby',
           'tomBoxHideRight', 'tomScreen', 'tomTextFlowWN',
           'tomDoubleWave', 'tomRubyAlign010', 'tomCollapseStart',
           'tomListNumberAsUCLetter', 'tomTibetan', 'tomAlignLeft',
           'tomMathIntraSpace', 'tomCharRepFromLcid',
           'tomFirstPageHeaderStory', 'tomMarchingBlackAnts',
           'tomEndnotesStory', 'tomLineSpace1pt5', 'tomAnimationMax',
           'MSANS', 'tomSelOvertype', 'tomTurkish', 'tomNoLink',
           'tomImageTypeMask', 'tomMathCollapseSel', 'tomFindStory',
           'MROMN', 'tomRubyBelow', 'tomMathWrapIndent',
           'tomSmallCaps', 'tomFontWeightExtraBlack', 'tomEnd',
           'tomMathInsRowAfter', 'tomMathAlphabetics', 'tomSentence',
           'tomGravityUI', 'tomMathPreSpace', 'tomReadOnly',
           'tomMathRelSize', 'tomConvertCRtoLF', 'tomCharFormat',
           'tomVai', 'tomMathDocSbSpOpUnchanged', 'tomBoxAlignCenter',
           'tomGlagolitic', 'tomFontWeightBlack', 'tomAutoBackColor',
           'tomAllowOffClient', 'tomTextize', 'tomDecSize',
           'tomEqArrayAlignBottomRow', 'tomInlineObjectArg',
           'tomToggleCase', 'tomDevanagari', 'tomParaEffectKeepNext',
           'tomParaFormat', 'tomFontStretchCondensed', 'tomResume',
           'ITextStory', 'tomRowUpdate', 'tomNoBreak',
           'tomStoryInactive', 'tomMathBrkBinDup', 'tomArmenian',
           'tomTifinagh', 'tomGravityIn', 'tomTextFlowMask',
           'tomMathArabicAlphabetics', 'tomHContCell',
           'tomTextFlowES', 'tomSparkleText', 'tomBaltic', 'tomStory',
           'MMONO', 'tomMarchingRedAnts', 'tomSymbol', 'tomStrikeout',
           'tomNeedTermOp', 'tomThick', 'tomMatchPattern',
           'tomExtend', 'tomMathInsRowBefore', 'tomDisabled',
           'tomListNone', 'MTAIL', 'tomSentenceCase', 'tomCyrillic',
           'tomConvertMathChar', 'tomTextFrameStory', 'tomStart',
           'tomNary', 'tomWipeDown', 'tomRadical', 'tomSubscript',
           'tomListNumberAsLCLetter', 'tomItalic', 'tomUsePoints',
           'tomListNumberedHindiAlpha', 'tomMathBuildDownOutermost',
           'tomShowULimPlaceHldr', 'tomCompressMax', 'tomArabic',
           'tomSuspend', 'tomAlignJustify', 'tomTrue',
           'tomImageRotate90', 'tomAlignWithTrailSpace',
           'tomStyleDefault', 'tomPhantomHorz', 'tomHair',
           'tomTextFlowSW', 'MSCRP', 'tomPage', 'tomDashes',
           'tomOverbar', 'tomFontWeightHeavy', 'tomLimitAlignCenter',
           'tomFirstPageFooterStory', 'tomWrapTextAround',
           'tomMathManualBreakMask', 'tomBengali', 'tomDots',
           'tomRubyAlignRight', 'tomPhantomHSmash',
           'tomEqArrayAlignTopRow', 'ITextPara', 'tomFontWeightBold',
           'tomListPeriod', 'tomLineSpaceDouble', 'tomDashDot',
           'tomFontWeightMax', 'tomStretchCharBelow', 'tomEmoji',
           'tomMathObjShadingStart', 'tomListNumberAsUCRoman',
           'tomBoxHideTop', 'tomSelectionColumn', 'tomInlineImage',
           'tomIncludeNumbering', 'tomKhmer', 'tomClientLink',
           'tomModWidthPairs', 'tomMathAlignBreakCenter',
           'tomNoWrapSides', 'tomAllowMathBold', 'tomHaveDelimiter',
           'tomMathDispAlignLeft', 'tomDoublestrike', 'tomShadow',
           'tomSpecialChar', 'tomDontSelectText',
           'tomLimitAlignRight', 'tomNormalCaret', 'tomMathCFCheck',
           'tomStretchCharAbove', 'MITAL', 'tomKannada',
           'tomRowApplyDefault', 'tomOsmanya',
           'tomListNumberAsArabic', 'tomTranslateTableCell',
           'tomNoHidden', 'tomLasVegasLights',
           'tomTransparentForPositioning', 'OBJECTTYPE',
           'tomRubyAlign121', 'tomMathEnter', 'tomEthiopic',
           'tomStretchStack', 'tomMathObjShadingEnd', 'ITextDocument',
           'tomUseXamlRect', 'tomConvertRTF', 'tomAlignBar',
           'ISequentialStream', 'tomMathSubscript', 'tomUseTwips',
           'tomLineSpaceExactly', 'MANCODE', 'tomSelRange',
           'tomFoldMathAlpha', 'tomMatchMathFont', 'tomDefaultTab',
           'tomUpperCase', 'tomListNumberedBlackCircleWingding',
           'tomGetHeightOnly', 'tomParaEffectTableRowDelimiter',
           'tomMathBrkBinSubMM', 'tomHStartCell', 'tomSelReplace',
           'tomMathSingleChar', 'tomStoryActiveDisplayUI',
           'tomMathDispDef', 'tomSuperscript', 'tomMathBrkBinSubMP',
           'tomListNumberedJpnChS', 'tomParaPropMathAlign', 'tomWord',
           'tomMathWrapRight', 'tomIgnoreCurrentFont',
           'tomWordDocument', 'tomTaiLe', 'MMATH', 'tomGurmukhi',
           'tomSelectionFrame', 'tomAlignInterLetter',
           'tomShareDenyRead', 'tomMatrixAlignMask', 'tomSinhala',
           'tomEquationArray', 'tomVLowCell', 'tomEastEurope',
           'tomTelugu', 'tomSelStartActive',
           'tomFontAlignmentBaseline', 'tomEnclose', 'tomGB2312',
           'tomOpChar', 'tomSizeScript', 'tomUndefined',
           'tomGeorgian', 'tomPhagsPa', 'MLOOP', 'tomRubyAlignLeft',
           'tomThai', 'tomLineSpacePercent', 'tomPhantomZeroDescent',
           'MOPENA', 'tomFunctionTypeIsLim', 'tomDocMathBuild',
           'tomStretchBaseAbove', 'tomCompressPunctuationAndKana',
           'tomIncIncSize', 'tomColumn', 'tomFontStretchDefault',
           'tomBrackets', 'tomListNumberedHindiNum',
           'tomEvenPagesFooterStory', 'tomBlinkingBackground',
           'tomEqArrayLayoutWidth', 'tomTeX', 'tomMathZoneDisplay',
           'tomFontAlignmentBottom', 'tomSubSup',
           'tomFontWeightExtraBold', 'tomRoundedBoxDashStyleMask',
           'tomReplaceStory', 'ITextStoryRanges', 'tomTabBack',
           'tomUnderlineTrailSpace', 'tomBraille', 'tomNoSelection',
           'tomMathAlignBreakLeft', 'tomMathParaAlignLeft',
           'tomMathDocDiffUpright', 'ITextStrings',
           'tomMathMakeSubSup', 'tomHebrew', 'tomNoMathZoneBrackets',
           'tomSizeScriptScript', 'tomUnderlinePositionAbove',
           'tomRE10Mode', 'tomEvenPagesHeaderStory',
           'tomMathInterSpace', 'tomListNumberedArabic2',
           'tomListNumberedArabic1', 'tomEllipsisPresent',
           'tomAlignNewspaper', 'ITextSelection2', 'tomSelActive',
           'tomThickLongDash', 'tomGrowWithContent',
           'tomInlineObjectStart', 'tomNoUpScroll',
           'tomMathDeleteRow', 'tomMongolian', 'tomEmbeddedFont',
           'tomEllipsisMode', 'tomParaEffectRTL',
           'tomParaEffectCollapsed', 'tomMathDocDiffItalic',
           'tomKayahli', 'tomDisableSmartFont',
           'tomMathDispAlignCenterGroup', 'tomCharset',
           'tomCharRepMax', 'tomMathAutoComplete', 'MSTRCH',
           'tagSTATSTG', 'tomMathDocDiffDefault',
           'tomMathMakeLeftSubSup', 'tomText', 'tomPhantomVert',
           'tomSelectionNormal', 'tomStyleScriptScript',
           'tomMathBrkBinSubMask', 'tomFootnotesStory',
           'tomAutoSpaceAlpha', 'tomMatchWord',
           'tomMatrixAlignCenter', 'tomMathAutoCorrect',
           'tomGravityBack', 'tomScratchStory', 'tomObjectArg',
           'tomTableColumn', 'tomMathDispAlignMask',
           'tomLimitsDefault', 'tomMove', 'tomMathMakeFracSlashed',
           'tomConvertMathML', 'ITextStoryRanges2', 'ITextFont2',
           'tomFontAlignmentMax', 'ITextFont', 'tomDefault',
           'tomUsymbol', 'tomListNumberedWhiteCircleWingding',
           'tomMathDispNaryGrow', 'tomMathBrkBinBefore',
           'tomStoryActiveUI', 'IStream', 'tomParaEffectNoLineNumber',
           'tomMathParaAlignCenterGroup', 'MOPEN',
           'tomListNumberAsLCRoman', 'tomMathDeleteCol',
           'tomCreateAlways', 'tomListNumberedArabicWide',
           'tomAllCaps', 'tomLimbu', 'tomEquals', 'tomNoAnimation',
           'tomImageRotate180', 'tomMathAlignBreakRight',
           'tomMathBuildUpRecurse', 'tomSelAtEOL', 'tomImageRotate0',
           'tomEllipsisNone', 'tomChemicalFormula', 'tomListMinus',
           'tomOpenExisting', 'tomLimitsOpposite', 'tomEq', 'MINIT',
           'tomOriya', 'tomFontWeightLight', 'tomSlashedFraction',
           'tomAutoTextColor', 'tomStyleScript', 'tomAccent',
           'tomNewTaiLue', 'tomVietnamese', 'tomFriendlyLinkAddress',
           'tomRowHeightActual', 'tomLanguageTag',
           'tomMathDocDiffOpenItalic', 'MFRAK', 'tomFraction',
           'tomWords', 'tomToggle', 'tomLowerLimit',
           'tomInlineObject', 'tomMathEnableRtl', 'tomFontStretch',
           'tomStoryLength', 'tomBoxHideLeft',
           'tomFunctionTypeTakesArg', 'tomMathDispAlignCenter',
           'tomSpaceBinary', 'tomIgnoreNumberStyle',
           'tomFontStretchSemiExpanded', 'tomTransparentForSpacing',
           'tomOutline', 'tomUnderlinePositionAuto', 'ITextDisplays',
           'tomUpperLimit', 'tomDefaultCharRep', 'tomMathBackspace',
           'tomTransform', 'tomAutoLinkURL', 'tomFontAlignmentAuto',
           'tomIgnoreTrailSpacing', 'tomStoryActiveDisplay', 'tomOEM',
           'tomFontBound', 'tomCanCopy', 'MBOLD', 'tomPhantomDSmash',
           'tomLimitAlignMask', 'tomSimpleText', 'tomTamil',
           'tomMathMakeFracLinear', 'tomGravityOut', 'tomUnlink',
           'tomListPlain', 'tomFalse', 'tomSpaceOrd',
           'tomFontWeightSemiBold', 'tomFontAlignmentTop',
           'tomFontStretchUltraCondensed', 'tomStyleDisplay',
           'tomUnicodeBiDi', 'tomFontWeightExtraLight',
           'tomSpaceSkip', 'tomApplyTmp', 'tomMathDocEmptyArgAlways',
           'tomMathChangeMask', 'tomBackward', 'tomKoreanBlockCaret',
           'tomParaEffectBox', 'tomNoUCGreekItalic', 'tomLink',
           'tomPhantom', 'tomFontPropAlign', 'tomLine',
           'tomProcessId', 'tomMathDispNarySubSup', 'tomParagraph',
           'tomStyleScriptCramped', 'tomAlignMatchAscentDescent',
           'tomAlignRight', 'tomThickDash', 'tomMathRemoveOutermost',
           'tomModWidthSpace', 'tomFontStretchNormal', 'tomLeafLine',
           'tomNoIME', 'tomMatchAscii', 'tomTabHere',
           'tomRoundedBoxCapStyleMask', 'tomSuperscriptCF',
           'tomMathDocEmptyArgNever', 'tomRunic', 'tomRow',
           'tomSizeText', 'tomSpaces', 'tomFontStretchExpanded',
           'tomMathVariant', 'ITextRow', 'tomFontWeightRegular',
           'tomMathArgShadingStart', 'tomExtendedChar',
           'tomFontStyle', 'tomMatchFontSignature', 'tomJamo',
           'tomStyleTextCramped', 'tomGothic', 'tomCacheParms',
           'tomUnhide', 'tomGravityFore', 'tomAutoLinkPhone',
           'ITextRange2', 'tomApplyLater', 'tomHorzVert',
           'tomMathBrkBinMask', 'tomFontStretchSemiCondensed',
           'tomBold', 'tomOpenAlways', 'tomOgham',
           'tomEqArrayAlignCenter', 'tomFontWeightDefault',
           'tomWindow', 'tomThickDashDot', 'tomGravityForward',
           'tomNone', 'tomAtEnd', 'tomShareDenyWrite', 'tomMathTab',
           'tomLao', 'MISOL', 'tomObject', 'tomMathEqAlign',
           'tomListNoNumber', 'tomBox', 'tomSelectionInlineShape',
           'tomMathZone', 'tomApplyRtfDocProps',
           'tomMathAutoCorrectExt', 'tomMatrixAlignBottomRow',
           'tomThaana', 'tomUnderlinePositionMax', 'ITextDocument2',
           'tomLimitsUnderOver', 'tomLayoutColumn',
           'tomParaStyleHeading7', 'tomParaStyleHeading6',
           'tomParaStyleHeading5', 'tomParaStyleHeading4',
           'tomParaStyleHeading3', 'tomParaStyleHeading2',
           'tomParaStyleHeading1', 'tomAlignDecimal', 'tomCharacter',
           'tomParaStyleHeading9', 'tomParaStyleHeading8',
           'tomAboriginal', 'tomLowerCase', 'tomSelectionRow',
           'tomShimmer', 'tomMath', 'tomParaEffectKeep', 'tomDotted',
           'tomKharoshthi', 'tomCherokee',
           'tomParaEffectOutlineLevel', 'tomMalayalam',
           'tomMathMakeFracStacked', 'tomAlignScaled',
           'tomMathZoneSurround', 'tomStyleText', 'tomUndoLimit',
           'ITextSelection', 'tomThickDashDotDot', 'tomBoxStrikeV',
           'tomMainTextStory', 'tomStyleDisplayCramped', 'tomTable',
           'tomHTML', 'tomListNumberedJpnKor', 'tomMathInsColAfter',
           'tomBoxStrikeH', 'tomLineSpaceSingle', 'tomGreek',
           'tomSubscriptCF', 'tomRoundedBoxNullRadius',
           'tomRoundedBoxCompact', 'tomImageRotate270',
           'tomFontStretchUltraExpanded', 'tomParaStyleNormal',
           'tomDeseret', 'tomShowMatPlaceHldr',
           'tomMathBuildUpArgOrZone', 'tomFontStretchExtraCondensed',
           'tomFunctionTypeTakesLim2', 'tomIncSize', 'tomMathLMargin',
           'tomProtected', 'tomMathBreakCenter', 'tomCollapseEnd',
           'tomSpaceDifferential', 'tomFontStyleUpright',
           'tomAutoSpaceNumeric', 'tomUseCRLF', 'tomAutoLinkEmail',
           'tomAnsi', 'tomListParentheses', 'tomMatchCase',
           'tomFontAlignmentCenter', 'tomSubSupAlign',
           'tomMathBreakRight', 'tomSylotiNagri',
           'tomAutoSpaceParens', 'tomSelectionShape', 'tomRTF',
           'tomTitleCase', 'tomMathAutoCorrectOpPairs',
           'tomAutoLinkPath', 'tomSpaceDefault', 'tomConstants',
           'tomUseAtFont', 'tomBoxedFormula', 'tomFontWeightNormal',
           'tomFontStyleItalic', 'tomGetTextForSpell', 'tomNKo',
           'tomSingle', 'tomListNumberedThaiNum', 'tomThickDotted',
           'tomCell', 'tomMathApplyTemplate', 'tomParaEffectTable',
           'tomMathDocEmptyArgAuto', 'tomTextFlowNE', 'tomRevised',
           'tomBoxHideBottom', 'tomEllipsisState',
           'tomPrimaryHeaderStory', 'tomLimitAlignLeft',
           'tomHardParagraph', 'tomBIG5', 'tomPhantomZeroWidth',
           'tomParaEffectPageBreakBefore', 'tomTabNext',
           'tomMathDispAlignRight', 'tomPhantomZeroAscent',
           'tomMathDispIntUnderOver', 'tomLines', 'tomAlignInterWord',
           'tomClientCoord', 'tomGravityBackward',
           'tomMathZoneNoBuildUp', 'tomForward', 'tomUnderline',
           'tomCluster', 'tomListBullet', 'tomConvertLinearFormat',
           'tomEnableSmartFont', 'tomParaEffectSideBySide',
           'tomMathDispFracTeX',
           '__MIDL___MIDL_itf_tom_0000_0000_0001', 'tomBoxStrikeTLBR',
           '__MIDL___MIDL_itf_tom_0000_0000_0003',
           '__MIDL___MIDL_itf_tom_0000_0000_0002', 'ITextRange',
           'tomListNumberedHindiAlpha1', 'tomListNumberedCircle',
           'tomHangul', 'tomFunctionTypeNone', 'tomThickLines',
           'tomPhantomShow', 'tomAutoColor', 'tomSelfIME',
           'tomHorzHorz', 'tomFontStretchExtraExpanded',
           'tomMathDeleteArg', 'tomMatrix', 'tomImageFlipH',
           'tomImageFlipV', 'tomDouble', 'tomFunctionApply',
           'tomMathDocDiffMask', 'tomCanRedo', 'tomUnderbar',
           'tomUnknownStory', 'tomEqArrayAlignMask',
           'tomMathBrkBinAfter', 'tomSyriac', 'tomMatchCharRep',
           'tomFunctionTypeTakesLim', 'tomNullCaret',
           'tomEllipsisEnd', 'tomYi', 'tomMathParaAlignRight',
           'tomPasteFile', 'tomConvertOMML', 'tomCommentsStory',
           'tomOverlapping', 'tomMathSuperscript',
           'tomMathParaAlignDefault', 'tomStyleScriptScriptCramped',
           'tomDash', 'tomCheckTextLimit', 'tomRubyAlignCenter',
           'tomParaEffectNoWidowControl']
from comtypes import _check_version; _check_version('')
