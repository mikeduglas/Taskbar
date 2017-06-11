  INCLUDE('svcomdef.inc'),ONCE

  COMPILE ('=== DO LINK', _ecgOleAutLinkMode_)
OleAutClass           CLASS,TYPE,LINK('oleautcg.lib', _APILinkMode_)
                      END
 === DO LINK

DoVerb                ITEMIZE,PRE(DOVERB)
Primary                 EQUATE(0)
Show                    EQUATE(-1)
Open                    EQUATE(-2)
Hide                    EQUATE(-3)
UIActivate              EQUATE(-4)
InPlaceActivate         EQUATE(-5)
DiscardUndoState        EQUATE(-6)
Properties              EQUATE(-7)
                      END

  SECTION('OleWrapper')

  
OleWrapperClass       CLASS, TYPE, MODULE('easyole.clw'), LINK('easyole.clw')
IsLibInitialized        BOOL,PROTECTED
OleCtrl                 SIGNED  !,PROTECTED
ParentCtrl              SIGNED,PROTECTED
rclsid                  REFCLSID,PROTECTED
lpInterface             LONG,PROTECTED
ClassName               CSTRING(51),PROTECTED
debug                   BOOL,PROTECTED
LicenseKey              STRING(256)

Construct               PROCEDURE()
Destruct                PROCEDURE()
Init                    PROCEDURE(STRING ClassName,REFCLSID rclsid,<SIGNED OleCtrl>,LONG x,LONG y,LONG w, LONG h,SHORT doverb=0,BOOL debug=1),SIGNED,PROC,VIRTUAL
Init                    PROCEDURE(STRING ClassName,REFCLSID rclsid,<SIGNED OleCtrl>,<SIGNED Ctrl>,SHORT doverb=0,BOOL debug=1),SIGNED,PROC,VIRTUAL
Kill                    PROCEDURE(),VIRTUAL
GetInterface            PROCEDURE(),LONG
Resize                  PROCEDURE(LONG x,LONG y,LONG w, LONG h)
Resize                  PROCEDURE(<SIGNED Ctrl>)
ShowError               PROCEDURE(STRING ProcName,HRESULT HR)
Show                    PROCEDURE()
Hide                    PROCEDURE()
Dec2Hex                 PROCEDURE(ULONG pDec,BYTE pAsString=0),STRING,PROTECTED
                      END

  SECTION('OleWrapperEx')

OleWrapperExClass   CLASS(OleWrapperClass),TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
Construct             PROCEDURE(),REPLACE
Init                  PROCEDURE(STRING ProgId,STRING ClassName,SIGNED Ctrl),SIGNED,PROC
                    END

  SECTION('OlePicture')

OlePictureClass       CLASS,TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
Construct               PROCEDURE()
Destruct                PROCEDURE(),VIRTUAL

Get_lpPicture           PROCEDURE(STRING pPictureName),LONG
Get_Picture             PROCEDURE(STRING pPictureName),*Picture
Get_IPicture            PROCEDURE(STRING pPictureName),*IPicture
Get_IPictureDisp        PROCEDURE(STRING pPictureName),*IPictureDisp
Get_Address             PROCEDURE(STRING pPictureName,LONG riid),LONG,PRIVATE
                      END

  SECTION('OleFont')

OleFontClass          CLASS,TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
pIFontDisp              &IFontDisp,PROTECTED

Construct               PROCEDURE()
Destruct                PROCEDURE(),VIRTUAL

Init                    PROCEDURE(LONG lpIFontDisp)

get_Name                PROCEDURE(),STRING
put_Name                PROCEDURE(STRING pName)
get_Size                PROCEDURE(),REAL  !CY
put_Size                PROCEDURE(REAL pSize)
get_Bold                PROCEDURE(),BOOL
put_Bold                PROCEDURE(BOOL bold)
get_Italic              PROCEDURE(),BOOL
put_Italic              PROCEDURE(BOOL italic)
get_Underline           PROCEDURE(),BOOL
put_Underline           PROCEDURE(BOOL underline)
get_Strikethrough       PROCEDURE(),BOOL
put_Strikethrough       PROCEDURE(BOOL strikethrough)
get_Weight              PROCEDURE(),SHORT
put_Weight              PROCEDURE(SHORT weight)
get_Charset             PROCEDURE(),SHORT
put_Charset             PROCEDURE(SHORT charset)

getProp                 PROCEDURE(LONG pDispId),*?,PRIVATE
putProp                 PROCEDURE(LONG pDispId,? pValue,LONG pVtType),PRIVATE
                      END

  SECTION('OleCollection')

OleCollectionClass    CLASS,TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
m_Disp                 &IDispatch,PROTECTED
m_Enum                 &IEnumVARIANT,PROTECTED
rgVar                   VARIANT,PROTECTED

Construct               PROCEDURE()
Destruct                PROCEDURE(),VIRTUAL

Init                    PROCEDURE(LONG pICollection),HRESULT,PROC,VIRTUAL
Init                    PROCEDURE(*IDispatch pICollection),HRESULT,PROC,VIRTUAL
Kill                    PROCEDURE(),VIRTUAL

ForEach                 PROCEDURE(*LONG pItem),HRESULT,VIRTUAL  !-- returns an address of IItem interface (use oItem.Init(pItem))
Count                   PROCEDURE(),LONG,VIRTUAL

GetItem                 PROCEDURE(? pIndex), *OleItemClass      !-- [in] long/string
                      END

!-- represents an item of a collection
OleItemClass        CLASS,TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
m_Disp                &IDispatch,PROTECTED

Construct             PROCEDURE()
Destruct              PROCEDURE(),VIRTUAL

Init                  PROCEDURE(LONG ppItem),HRESULT,PROC,VIRTUAL
Init                  PROCEDURE(*IDispatch pIItem),HRESULT,PROC,VIRTUAL
Kill                  PROCEDURE(),VIRTUAL

GetValue              PROCEDURE(STRING pProp), ?
GetValue              PROCEDURE(), ?              !-- 'Value' prop
SetValue              PROCEDURE(STRING pProp, STRING pValue)
SetValue              PROCEDURE(STRING pValue)    !-- 'Value' prop
                    END

  SECTION('OleSafeArray')

!SafeArray manipulations

OleSafeArrayClass     CLASS,TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
vt                      LONG, PROTECTED

lpSA                    LONG,PROTECTED              !the address of the array
created                 BYTE,PROTECTED              !created, must be destroyed by SafeArrayDestroy()

Construct               PROCEDURE()
Destruct                PROCEDURE(),VIRTUAL

Init                    PROCEDURE(*VARIANT pVar),HRESULT,PROC,VIRTUAL  !get copy of existing array; S_OK if SAFEARRAY
Init                    PROCEDURE(VARTYPE pVTType,LONG pElements),HRESULT,PROC,VIRTUAL  !create an array with dim=1; S_OK if SAFEARRAY
Init                    PROCEDURE(VARTYPE pVTType,LONG nDims,LONG[] pElements),HRESULT,PROC,VIRTUAL  !create an array with; S_OK if SAFEARRAY

Kill                    PROCEDURE(),VIRTUAL
!Destroy                 PROCEDURE(),VIRTUAL

GetVartype              PROCEDURE(),VARTYPE,VIRTUAL
GetElemsize             PROCEDURE(),UNSIGNED,VIRTUAL
GetNDims                PROCEDURE(),LONG,VIRTUAL
GetLBound               PROCEDURE(UNSIGNED pDim),LONG,VIRTUAL
GetUBound               PROCEDURE(UNSIGNED pDim),LONG,VIRTUAL
GetNElements            PROCEDURE(UNSIGNED pDim),LONG,VIRTUAL
GetArray                PROCEDURE(),LONG,VIRTUAL
GetVariant              PROCEDURE(*VARIANT pv),VIRTUAL

GetElement              PROCEDURE(LONG pDim,LONG pIndex,*VARIANT pElement),HRESULT,PROC,VIRTUAL
PutElement              PROCEDURE(LONG pDim,LONG pIndex,LONG pElement),HRESULT,PROC,VIRTUAL
                      END

OleSafeArrayExClass   CLASS(OleSafeArrayClass),TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
!vt                      LONG, PROTECTED
!
Init                    PROCEDURE(VARTYPE pVTType,LONG pElements, LONG pLBound), HRESULT, PROC              !create an array with dim=1
Init                    PROCEDURE(VARTYPE pVTType,LONG pDims, LONG pElements, LONG pLBound), HRESULT, PROC  !create an array with dim=pDims

!-- Init multi-dimensional safearray (cDims = FIELDS(q)) with vt=VT_VARIANT and copies queue
InitQ                   PROCEDURE(QUEUE q, LONG pLBound = 0), HRESULT, PROC
!-- Init 1-dimensional safearray with vt=VT_VARIANT and copies queue
InitQ1                  PROCEDURE(QUEUE q, LONG pLBound = 0), HRESULT, PROC

InitAsPPSA              PROCEDURE(LONG ppSA), HRESULT, PROC  ! ppSA = SAFEARRAY**

PutVariant              PROCEDURE(LONG pDim,LONG pIndex,*VARIANT pVar,LONG vt=VT_EMPTY),HRESULT,PROC
ReDim                   PROCEDURE(LONG newSize),HRESULT,PROC       ! This function changes dimension of a 1-dimensioned safearray.
CopyToVariant           PROCEDURE(),*VARIANT
CopyToVariant           PROCEDURE(*gVARIANT pgv)
Destroy                 PROCEDURE(LONG psa),HRESULT,PROC

GetData                 PROCEDURE(LONG pBuffer, LONG pBufferSize),HRESULT,PROC
PutData                 PROCEDURE(LONG pBuffer, LONG pBufferSize),HRESULT,PROC
                      END

  SECTION('OleVariant')

!Variant manipulations

OleVariantClass       CLASS,TYPE,MODULE('easyole.clw'),LINK('easyole.clw')
GetValue                PROCEDURE(? par),*VARIANT
Missing                 PROCEDURE(),*VARIANT    !missing argument, vt_error
GetAddress              PROCEDURE(? par),LONG
                      END