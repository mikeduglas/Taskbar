  MEMBER
  include('svcom.inc'),ONCE
  MAP
    include('svapifnc.inc'),ONCE

    MODULE('oleaut32.dll')
      ecg_CreateStreamOnHGlobal(LONG hGlobal,BOOL fDeleteOnRelease,LONG ppstm),LONG,PASCAL,PROC,NAME('CreateStreamOnHGlobal')
      ecg_OleLoadPicture(*IStream pStream,LONG lSize,BOOL fRunmode,LONG riid,LONG ppvObj),LONG,PASCAL,PROC,NAME('OleLoadPicture')

      ecg_SafeArrayCreate(VARTYPE vt,UNSIGNED cDims,LONG rgsabound),LONG,PASCAL,NAME('SafeArrayCreate')
      ecg_SafeArrayGetVartype(LONG psa,*VARTYPE vt),HRESULT,PASCAL,NAME('SafeArrayGetVartype')
      ecg_SafeArrayGetDim(LONG psa),UNSIGNED,PASCAL,NAME('SafeArrayGetDim')
      ecg_SafeArrayGetElemsize(LONG psa),UNSIGNED,PASCAL,NAME('SafeArrayGetElemsize')
      ecg_SafeArrayGetLBound(LONG psa,UNSIGNED nDim,*LONG plLbound),HRESULT,PASCAL,PROC,NAME('SafeArrayGetLBound')
      ecg_SafeArrayGetUBound(LONG psa,UNSIGNED nDim,*LONG plUbound),HRESULT,PASCAL,PROC,NAME('SafeArrayGetUBound')
      ecg_SafeArrayDestroy(LONG psa),HRESULT,PASCAL,PROC,NAME('SafeArrayDestroy')
      ecg_SafeArrayGetElement(LONG psa,LONG prgIndices,LONG pv),HRESULT,PASCAL,NAME('SafeArrayGetElement')
      ecg_SafeArrayPutElement(LONG psa,LONG prgIndices,LONG pv),HRESULT,PASCAL,NAME('SafeArrayPutElement')
      ecg_SafeArrayRedim(LONG psa,LONG rgsabound),HRESULT,PASCAL,NAME('SafeArrayRedim')
      ecg_SafeArrayCopy(LONG sa, *LONG psaOut),HRESULT,RAW,PASCAL,NAME('SafeArrayCopy'),PROC
   
      ecg_SafeArrayAccessData(LONG psa, *LONG pvData),HRESULT,RAW,PASCAL,NAME('SafeArrayAccessData')
      ecg_SafeArrayUnaccessData(LONG psa),HRESULT,RAW,PASCAL,PROC,NAME('SafeArrayUnaccessData')
    END
    MODULE('WINAPI')
      ecg_CreateFile(*CSTRING,ULONG,ULONG,LONG,ULONG,ULONG,UNSIGNED=0),UNSIGNED,RAW,PASCAL,NAME('CreateFileA')
      ecg_GetFileSize(UNSIGNED,*ULONG),ULONG,PASCAL,NAME('GetFileSize')
      ecg_ReadFile(UNSIGNED,LONG,ULONG,*ULONG,LONG),BOOL,PASCAL,RAW,NAME('ReadFile')
      ecg_CloseHandle(UNSIGNED),BOOL,PASCAL,PROC,NAME('CloseHandle')
      ecg_GlobalAlloc(LONG uFlags,LONG dwBytes),LONG,PASCAL,NAME('GlobalAlloc')
      ecg_GlobalFree(SIGNED hMem),SIGNED,PASCAL,PROC,NAME('GlobalFree')
      ecg_GlobalLock(SIGNED hMem),LONG,PASCAL,PROC,NAME('GlobalLock')
      ecg_GlobalUnlock(SIGNED hMem),BOOL,PASCAL,PROC,NAME('GlobalUnlock')
      ecg_memcpy(LONG lpDest,LONG lpSource,LONG nCount),LONG,PROC,NAME('_memcpy')
      ecg_VariantCopy(*VARIANT vargDest,*VARIANT vargSrc),HRESULT,PASCAL,PROC,NAME('VariantCopy')
      ecg_VariantCopyInd(*VARIANT vargDest,*VARIANT vargSrc),HRESULT,PASCAL,PROC,NAME('VariantCopyInd')
      ecg_VariantChangeType(*VARIANT vargDest,*VARIANT vargSrc, USHORT flags=0, VARTYPE vt),HRESULT,PASCAL,PROC,NAME('VariantChangeType')
      ecg_VariantClear(*VARIANT varg),HRESULT,PASCAL,PROC,NAME('VariantClear')

      ecg_lstrlen(LONG lpWideString),LONG,NAME('lstrlenw'),PASCAL   !Randy Rogers' fix

      ecg_CoInitializeEx(LONG,ULONG),HRESULT,RAW,PASCAL,NAME('CoInitializeEx')
      ecg_OutputDebugString(*CSTRING lpOutputString),PASCAL,RAW,NAME('OutputDebugStringA')
    END

    ecg::DebugInfo(STRING s)

    INCLUDE('getvartype.inc', 'DECLARATIONS')
  END
  INCLUDE('ecom2inc.def'),ONCE
  INCLUDE('easyole.inc')
  INCLUDE('getvartype.inc', 'CODE')

ecg::DebugInfo                PROCEDURE(STRING s)
pfx                             STRING('[EasyOle] ')
cs                              CSTRING(LEN(s) + LEN(pfx) + 1)
  CODE
  cs = pfx & s
  ecg_OutputDebugString(cs)

!============= OleWrapperClass ======================
OleWrapperClass.Construct  PROCEDURE()
  CODE
  SELF.IsLibInitialized=CHOOSE(INLIST(CoInitialize(0),S_OK,S_FALSE)=0,false,true)
  !SELF.IsLibInitialized=CHOOSE(INLIST(ecg_CoInitializeEx(0,COINIT_APARTMENTTHREADED),S_OK,S_FALSE)=0,false,true)
  !SELF.IsLibInitialized=CHOOSE(INLIST(ecg_CoInitializeEx(0,COINIT_MULTITHREADED),S_OK,S_FALSE)=0,false,true)

OleWrapperClass.Destruct   PROCEDURE()
  CODE
  CoUnInitialize()

OleWrapperClass.Init      PROCEDURE(STRING ClassName,REFCLSID rclsid,<SIGNED OleCtrl>,LONG x,LONG y,LONG w, LONG h,SHORT doverb,BOOL debug)
WideProgId                STRING(256)
CWideProgId               &CWideStr
CstrProgId                &Cstr
ProgId                    CSTRING(128)
lpOleStr                  LONG
lpInterface               LONG
cobj                      CSTRING(20)
HR                        HRESULT(S_FALSE)

ErrLocation               CSTRING(101)

lenWideProgId             LONG  !Randy Rogers' fix

  CODE
  SELF.debug=debug

  IF NOT SELF.IsLibInitialized
    IF SELF.debug
      SELF.ShowError('COM initializing',S_FALSE)
    END
    RETURN 0
  END

  SELF.ClassName=CLIP(ClassName)
  SELF.rclsid=rclsid

  ASSERT(SELF.rclsid)
  IF SELF.rclsid
    IF OMITTED(4)
      SELF.OleCtrl=CREATE(0,CREATE:OLE)
    ELSE
      SELF.OleCtrl=OleCtrl
    END
    HR=ProgIDFromCLSID(SELF.rclsid,lpOleStr)
    IF HR=S_OK
      !PEEK(lpOleStr,WideProgId)
      !-- Randy Rogers' fix
      lenWideProgId = (ecg_lstrlen(lpOleStr)+1)*2
      ASSERT(lenWideProgId <= 256)
      ecg_memcpy(ADDRESS(WideProgId),lpOleStr,lenWideProgId)
      !--

      CWideProgId &= NEW(CWideStr)
      IF CWideProgId.Init(ADDRESS(WideProgId))
        CstrProgId &= NEW(CStr)
        IF CstrProgId.Init(CWideProgId)
          ProgId=CstrProgId.GetCstr()
          !-- <start of> Carlos Gutierrez
          IF CLIP(SELF.LicenseKey) <> ''
            SELF.OleCtrl{prop:create}=CLIP(ProgId) & '\!' & CLIP(SELF.LicenseKey)
          ELSE
            SELF.OleCtrl{prop:create}=ProgId
          END
          !-- <end of> Carlos Gutierrez

          IF doverb <= 0
            SELF.OleCtrl{prop:doverb}=doverb
          END
          cobj = SELF.OleCtrl{prop:object}
          !-- for registration free OLE see http://www.wasm.ru/forum/viewtopic.php?id=24387
          IF cobj[1] <> '`'
            HR=S_FALSE
            ErrLocation='bad OleCtrl{{prop:object}'
          ELSE
            SELF.lpInterface = cObj[2 : LEN(CLIP(cobj))]
            SELF.OleCtrl{prop:release}=cobj

            SELF.OleCtrl{prop:xpos}=x
            SELF.OleCtrl{prop:ypos}=y
            SELF.OleCtrl{prop:width}=w
            SELF.OleCtrl{prop:height}=h

            UNHIDE(SELF.OleCtrl)
          END
        ELSE
          HR=S_FALSE
          ErrLocation='CstrProgId.Init'
        END
        DISPOSE(CstrProgId)
      ELSE
        HR=S_FALSE
        ErrLocation='CWideProgId.Init'
      END
      DISPOSE(CWideProgId)
    ELSE
      ErrLocation='ProgIDFromCLSID'
    END
  ELSE
    ErrLocation='rclsid=0'
  END
  IF HR <> S_OK
    IF ~ErrLocation
      SELF.ShowError('Init',HR)
    ELSE
      SELF.ShowError('Init ('&ErrLocation&')',HR)
    END
    DESTROY(SELF.OleCtrl)
    SELF.OleCtrl=0
  END
  RETURN SELF.OleCtrl

OleWrapperClass.Init       PROCEDURE(STRING ClassName,REFCLSID rclsid,<SIGNED OleCtrl>,<SIGNED Ctrl>,SHORT doverb,BOOL debug)
  CODE
  IF OMITTED(5)
    IF ~OMITTED(4)
      SELF.Init(ClassName,rclsid,OleCtrl,OleCtrl{prop:xpos},OleCtrl{prop:ypos},OleCtrl{prop:width},OleCtrl{prop:height},doverb,debug)
    ELSE
      SELF.Init(ClassName,rclsid,,0,0,100,100,doverb,debug)
    END
  ELSE
    SELF.ParentCtrl=Ctrl
    IF ~OMITTED(4)
      SELF.Init(ClassName,rclsid,OleCtrl,Ctrl{prop:xpos},Ctrl{prop:ypos},Ctrl{prop:width},Ctrl{prop:height},doverb,debug)
    ELSE
      SELF.Init(ClassName,rclsid,,Ctrl{prop:xpos},Ctrl{prop:ypos},Ctrl{prop:width},Ctrl{prop:height},doverb,debug)
    END
  END
  RETURN SELF.OleCtrl

OleWrapperClass.Kill       PROCEDURE()
  CODE
  IF SELF.OleCtrl
    DESTROY(SELF.OleCtrl)
    SELF.OleCtrl=0
  END

OleWrapperClass.GetInterface PROCEDURE()
  CODE
  RETURN SELF.lpInterface

OleWrapperClass.Resize     PROCEDURE(LONG x,LONG y,LONG w, LONG h)
  CODE
  SETPOSITION(SELF.OleCtrl,x,y,w,h)

OleWrapperClass.Resize     PROCEDURE(<SIGNED Ctrl>)
  CODE
  IF OMITTED(2)
    SETPOSITION(SELF.OleCtrl,SELF.ParentCtrl{prop:xpos},SELF.ParentCtrl{prop:ypos},SELF.ParentCtrl{prop:width},SELF.ParentCtrl{prop:height})
  ELSE
    ASSERT(Ctrl>0)
    IF Ctrl>0
      SETPOSITION(SELF.OleCtrl,Ctrl{prop:xpos},Ctrl{prop:ypos},Ctrl{prop:width},Ctrl{prop:height})
    END
  END

OleWrapperClass.ShowError  PROCEDURE(STRING ProcName,HRESULT HR)
  CODE
  IF SELF.debug
    IF HR <> S_OK
      MESSAGE(CLIP(ProcName)&' failed ('&SELF.Dec2Hex(HR)&')',SELF.ClassName,ICON:EXCLAMATION)
    ELSE
      MESSAGE(CLIP(ProcName),SELF.ClassName,ICON:EXCLAMATION)
    END
  END

OleWrapperClass.Show  PROCEDURE()
  CODE
  IF SELF.OleCtrl AND SELF.OleCtrl{prop:hide}=true
    SELF.OleCtrl{prop:hide}=false
  END

OleWrapperClass.Hide  PROCEDURE()
  CODE
  IF SELF.OleCtrl AND SELF.OleCtrl{prop:hide}=false
    SELF.OleCtrl{prop:hide}=true
  END

OleWrapperClass.Dec2Hex            PROCEDURE(ULONG pDec,BYTE pAsString)
locHex    STRING(30)
  CODE
  LOOP UNTIL(~pDec)
    locHex = SUB('0123456789ABCDEF',1+pDec % 16,1) & CLIP(locHex)
    pDec = INT(pDec / 16)
  END

  IF pAsString  !'H' at the end and '0' at the start
    IF INRANGE(locHex[1],'A','F') THEN locHex='0' & clip(locHex) END
    locHex=CLIP(locHex) & 'H'
  END

  RETURN CLIP(locHex)

!============ OlePictureClass =======================
OlePictureClass.Construct           PROCEDURE()
  CODE

OlePictureClass.Destruct            PROCEDURE()
  CODE

OlePictureClass.Get_Address   PROCEDURE(STRING pPictureName,LONG riid)
pIStream                      &IStream
lpIStream                     LONG

lpIPicture                    LONG

szFile                        CSTRING(256)
hFile                         HANDLE
dwImageSize                   UNSIGNED
lpFileSizeHigh                ULONG
hGlobal                       HGLOBAL
pvData                        LONG
dwBytesRead                   ULONG
bRead                         BOOL
hr                            HRESULT

EC_INVALID_HANDLE_VALUE       EQUATE(-1)

  CODE
  !first get IStream from image file
  szFile=CLIP(pPictureName)
  hFile = ecg_CreateFile(szFile,GENERIC_READ,0,0,OPEN_EXISTING,0,0)
  IF hFile <> EC_INVALID_HANDLE_VALUE
    dwImageSize = ecg_GetFileSize(hFile,lpFileSizeHigh)
    IF dwImageSize > 0
      hGlobal = ecg_GlobalAlloc(GMEM_MOVEABLE,dwImageSize )
      IF hGlobal
        pvData = ecg_GlobalLock(hGlobal)
        IF pvData
          bRead = ecg_ReadFile(hFile,pvData,dwImageSize,dwBytesRead,0)
          ecg_GlobalUnlock(hGlobal)
          IF bRead
            hr=ecg_CreateStreamOnHGlobal(hGlobal,true,ADDRESS(lpIStream))
            IF hr=S_OK
              pIStream &= (lpIStream)
            END
          END
        END
      END
    END
    ecg_CloseHandle(hFile)
  END

  !then get IPicture or IPictureDisp
  IF ~pIStream &= NULL
    hr = ecg_OleLoadPicture(pIStream,dwImageSize,false,riid,ADDRESS(lpIPicture))
    pIStream.Release()
  END
  RETURN lpIPicture

OlePictureClass.Get_lpPicture       PROCEDURE(STRING pPictureName)
  CODE
  RETURN(SELF.Get_Address(pPictureName,ADDRESS(IID_IPictureDisp)))

OlePictureClass.Get_Picture         PROCEDURE(STRING pPictureName)
pPicture                            &Picture
  CODE
  pPicture &= (SELF.Get_Address(pPictureName,ADDRESS(IID_IPictureDisp)))
  RETURN pPicture

OlePictureClass.Get_IPicture        PROCEDURE(STRING pPictureName)
pIPicture                           &IPicture
  CODE
  pIPicture &= (SELF.Get_Address(pPictureName,ADDRESS(IID_IPicture)))
  RETURN pIPicture

OlePictureClass.Get_IPictureDisp    PROCEDURE(STRING pPictureName)
pIPictureDisp                       &IPictureDisp
  CODE
  pIPictureDisp &= (SELF.Get_Address(pPictureName,ADDRESS(IID_IPictureDisp)))
  RETURN pIPictureDisp


!============ OleFontClass =======================
OleFontClass.Construct          PROCEDURE()
  CODE

OleFontClass.Destruct           PROCEDURE()
  CODE

OleFontClass.Init               PROCEDURE(LONG lpIFontDisp)
  CODE
  ASSERT(lpIFontDisp <> 0)
  SELF.pIFontDisp &= (lpIFontDisp)

OleFontClass.get_Name           PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_NAME)

OleFontClass.put_Name           PROCEDURE(STRING pName)
bstrName                        BSTRING
  CODE
  bstrName=pName
  SELF.putProp(DISPID_FONT_NAME,bstrName,VT_BSTR)

OleFontClass.get_Size           PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_SIZE)

OleFontClass.put_Size           PROCEDURE(REAL pSize)
  CODE
  SELF.putProp(DISPID_FONT_SIZE,pSize,VT_CY)

OleFontClass.get_Bold           PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_BOLD)

OleFontClass.put_Bold           PROCEDURE(BOOL bold)
  CODE
  SELF.putProp(DISPID_FONT_BOLD,bold,VT_I4)

OleFontClass.get_Italic         PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_ITALIC)

OleFontClass.put_Italic         PROCEDURE(BOOL italic)
  CODE
  SELF.putProp(DISPID_FONT_ITALIC,italic,VT_I4)

OleFontClass.get_Underline      PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_UNDER)

OleFontClass.put_Underline      PROCEDURE(BOOL underline)
  CODE
  SELF.putProp(DISPID_FONT_UNDER,underline,VT_I4)

OleFontClass.get_Strikethrough  PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_STRIKE)

OleFontClass.put_Strikethrough  PROCEDURE(BOOL strikethrough)
  CODE
  SELF.putProp(DISPID_FONT_STRIKE,strikethrough,VT_I4)

OleFontClass.get_Weight         PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_WEIGHT)

OleFontClass.put_Weight         PROCEDURE(SHORT weight)
  CODE
  SELF.putProp(DISPID_FONT_WEIGHT,weight,VT_I2)

OleFontClass.get_Charset        PROCEDURE()
  CODE
  RETURN SELF.getProp(DISPID_FONT_CHARSET)

OleFontClass.put_Charset        PROCEDURE(SHORT charset)
  CODE
  SELF.putProp(DISPID_FONT_CHARSET,charset,VT_I2)

OleFontClass.getProp            PROCEDURE(LONG pDispId)
locDispParams                   LIKE(DISPPARAMS)
locVarResult                    VARIANT
locgVarResult                   LIKE(gVariant),OVER(locVarResult)
locExcepInfo                    LIKE(EXCEPINFO)
locuArgErr                      ULONG
locMethodName                   BSTRING
locdispid                       LONG
loccArgs                        UNSIGNED
HR                              HRESULT
RetVal                          ANY

  CODE
  IF ~SELF.pIFontDisp &= NULL
    locDispParams.cArgs=0
    HR=SELF.pIFontDisp.Invoke(pDispId,ADDRESS(IID_NULL),0,DISPATCH_PROPERTYGET,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
    IF HR = S_OK
      RetVal=locVarResult
    END
  END
  RETURN RetVal

OleFontClass.putProp            PROCEDURE(LONG pDispId,? pValue,LONG pVtType)
locDispParams                   LIKE(DISPPARAMS)
locVarResult                    VARIANT
locgVarResult                   LIKE(gVariant),OVER(locVarResult)
locExcepInfo                    LIKE(EXCEPINFO)
locuArgErr                      ULONG
locdispidNamed                  LONG(DISPID_PROPERTYPUT)
locpvargtype                    GROUP,DIM(1)
pvarg                             VARIANT
gpvarg                            LIKE(gVariant),OVER(pvarg)
                                END
HR                              HRESULT

  CODE
  IF ~SELF.pIFontDisp &= NULL
    locDispParams.cNamedArgs=1
    locDispParams.rgdispidNamedArgs=ADDRESS(locdispidNamed)
    locDispParams.rgvarg=ADDRESS(locpvargtype)
    CASE pVtType
    OF VT_I2
      locpvargtype[1].gpvarg.iVal=pValue
    OF VT_I4
      locpvargtype[1].gpvarg.lVal=pValue
    OF VT_CY
      !i know how to process only small positive currency values
      locpvargtype[1].gpvarg.cyVal.Lo=INT(pValue)*10000 + (pValue-INT(pValue))*10000
    ELSE
      locpvargtype[1].pvarg=pValue
    END
    locpvargtype[1].gpvarg.vt=pVtType
    locDispParams.cArgs=1
    HR=SELF.pIFontDisp.Invoke(pDispId,ADDRESS(IID_NULL),0,DISPATCH_PROPERTYPUT,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
  END

!============ OleCollectionClass =======================
OleCollectionClass.Construct    PROCEDURE()
  CODE

OleCollectionClass.Destruct     PROCEDURE()
  CODE
  SELF.Kill()

OleCollectionClass.Init         PROCEDURE(LONG pICollection)
pDisp                          &IDispatch
  CODE
  ASSERT(pICollection)
  IF ~pICollection
    RETURN S_FALSE
  END
  pDisp &= (pICollection)
  RETURN SELF.Init(pDisp)

OleCollectionClass.Init         PROCEDURE(*IDispatch pICollection)
locDispParams                   LIKE(DISPPARAMS)
locgVarResult                   LIKE(gVariant)
locVarResult                    VARIANT,OVER(locgVarResult)
locExcepInfo                    LIKE(EXCEPINFO)
locuArgErr                      ULONG
Unk                            &IUnknown
pUnk                            LONG
pEnum                           LONG
HR                              HRESULT
  CODE
  ASSERT(~pICollection &= NULL)
  IF pICollection &= NULL
    RETURN S_FALSE
  END

  IF ~SELF.m_Enum &= NULL
    SELF.Kill()
  END

  SELF.m_Disp &= pICollection

  locDispParams.cArgs=0
  HR=pICollection.Invoke(DISPID_NEWENUM,ADDRESS(IID_NULL),0,DISPATCH_METHOD+DISPATCH_PROPERTYGET,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
  IF HR <> S_OK
    RETURN HR
  END
  IF ~locgVarResult.lVal
    RETURN E_UNEXPECTED
  END
  pUnk=locgVarResult.lVal
  Unk &= (pUnk)
  HR=Unk.QueryInterface(ADDRESS(IID_IEnumVARIANT),pEnum)
  Unk.Release()
  IF HR <> S_OK
    RETURN E_NOINTERFACE
  END
  IF ~pEnum
    RETURN E_UNEXPECTED
  END
  SELF.m_Enum &= (pEnum)
  RETURN S_OK

OleCollectionClass.Kill         PROCEDURE()
  CODE
  IF ~SELF.m_Enum &= NULL
    SELF.m_Enum.Release()
    SELF.m_Enum &= NULL
  END
  IF ~SELF.m_Disp &= NULL
    SELF.m_Disp.Release()
    SELF.m_Disp &= NULL
  END
  CLEAR(SELF.rgVar)

OleCollectionClass.ForEach      PROCEDURE(*LONG pItem)
gvObj                           LIKE(gVariant),OVER(SELF.rgVar)
CeltFetched                     ULONG
HR                              HRESULT
  CODE
  pItem=0
  IF SELF.m_Enum &= NULL
    RETURN S_FALSE
  END
  HR=SELF.m_Enum.Next(1,SELF.rgVar,CeltFetched)
  IF HR=S_OK
    pItem=gvObj.lVal
  END
  RETURN HR

OleCollectionClass.Count        PROCEDURE()
pCount                          LONG
locDispParams                   LIKE(DISPPARAMS)
locVarResult                    VARIANT
locgVarResult                   LIKE(gVariant),OVER(locVarResult)
locExcepInfo                    LIKE(EXCEPINFO)
locuArgErr                      ULONG
locMethodName                   BSTRING
locdispid                       LONG
HR                              HRESULT
  CODE
  IF SELF.m_Enum &= NULL
    RETURN 0
  END
  locMethodName='Count'
  HR=SELF.m_Disp.GetIDsOfNames(ADDRESS(IID_NULL),ADDRESS(locMethodName),1,0,ADDRESS(locdispid))
  IF HR <> S_OK
    RETURN 0
  END
  locDispParams.cArgs=0
  HR=SELF.m_Disp.Invoke(locdispid,ADDRESS(IID_NULL),0,DISPATCH_METHOD+DISPATCH_PROPERTYGET,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
  IF HR <> S_OK
    RETURN 0
  END
  pCount=locVarResult
  RETURN pCount

OleCollectionClass.GetItem  PROCEDURE(? pIndex)
v:pIndex                      VARIANT
gv:pIndex                     LIKE(gVariant),OVER(v:pIndex)
locDispParams                 LIKE(DISPPARAMS)
locgVarResult                 LIKE(gVariant)
locVarResult                  VARIANT,OVER(locgVarResult)
locExcepInfo                  LIKE(EXCEPINFO)
locuArgErr                    ULONG
locMethodName                 BSTRING
locdispid                     LONG
loccArgs                      UNSIGNED
locpvargtype                  GROUP,DIM(1)
pvarg                           VARIANT
gpvarg                          LIKE(gVariant),OVER(pvarg)
                              END

oItem                         &OleItemClass

HR                            HRESULT

  CODE
  
  locMethodName='Item'
  HR=SELF.m_Disp.GetIDsOfNames(ADDRESS(IID_NULL),ADDRESS(locMethodName),1,0,ADDRESS(locdispid))
  IF HR <> S_OK
    MESSAGE('"Item" property not found, error '& HR)
    RETURN oItem
  END
  
  locDispParams.rgvarg=ADDRESS(locpvargtype)
  
  loccargs+=1
  locpvargtype[loccargs].pvarg=pIndex
  v:pIndex=pIndex
  locpvargtype[loccargs].gpvarg.vt=gv:pIndex.vt
  locDispParams.cArgs=loccargs
  
  HR=SELF.m_Disp.Invoke(locdispid,ADDRESS(IID_NULL),0,DISPATCH_METHOD+DISPATCH_PROPERTYGET,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
  IF HR < S_OK
    MESSAGE('Item() throws error '& HR)
    RETURN oItem
  END

  IF INLIST(locgVarResult.vt, VT_UNKNOWN, VT_DISPATCH)
    oItem &= NEW OleItemClass
    IF oItem.Init(locgVarResult.lval) <> S_OK
      DISPOSE(oItem)
      oItem &= NULL
    END
  END

  RETURN oItem

!============ OleItemClass =======================
OleItemClass.Construct  PROCEDURE()
  CODE

OleItemClass.Destruct   PROCEDURE()
  CODE
  SELF.Kill()

OleItemClass.Init   PROCEDURE(LONG ppItem)
pDisp                 &IDispatch
  CODE
  ASSERT(ppItem)
  IF ~ppItem
    RETURN S_FALSE
  END
  pDisp &= (ppItem)
  RETURN SELF.Init(pDisp)

OleItemClass.Init   PROCEDURE(*IDispatch pIItem)
  CODE
  ASSERT(~pIItem &= NULL)
  IF pIItem &= NULL
    RETURN S_FALSE
  END

  SELF.m_Disp &= pIItem
  RETURN S_OK

OleItemClass.Kill PROCEDURE()
  CODE
  IF ~SELF.m_Disp &= NULL
    SELF.m_Disp.Release()
    SELF.m_Disp &= NULL
  END

OleItemClass.GetValue   PROCEDURE(STRING pProp)
locDispParams             LIKE(DISPPARAMS)
locgVarResult             LIKE(gVariant)
locVarResult              VARIANT,OVER(locgVarResult)
locExcepInfo              LIKE(EXCEPINFO)
locuArgErr                ULONG
locMethodName             BSTRING
locdispid                 LONG
loccArgs                  UNSIGNED
LocRetVal                 HRESULT
HR                        HRESULT

  CODE
  ASSERT(NOT SELF.m_Disp &= NULL)
  IF SELF.m_Disp &= NULL 
    RETURN 0
  END

  locMethodName=pProp
  HR=SELF.m_Disp.GetIDsOfNames(ADDRESS(IID_NULL),ADDRESS(locMethodName),1,0,ADDRESS(locdispid))
  IF HR <> S_OK
    MESSAGE('"'& pProp &'" property not found, error '& HR)
    RETURN 0
  END

  locDispParams.cArgs=loccargs
  HR=SELF.m_Disp.Invoke(locdispid,ADDRESS(IID_NULL),0,DISPATCH_METHOD+DISPATCH_PROPERTYGET,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
  IF HR < S_OK
    MESSAGE('Get'& pProp &' throws error '& HR)
    RETURN 0
  END
  
  RETURN locVarResult

OleItemClass.GetValue   PROCEDURE()
  CODE
  RETURN SELF.GetValue('Value')

OleItemClass.SetValue   PROCEDURE(STRING pProp, STRING pValue)
locDispParams             LIKE(DISPPARAMS)
locgVarResult             LIKE(gVariant)
locVarResult              VARIANT,OVER(locgVarResult)
locExcepInfo              LIKE(EXCEPINFO)
locuArgErr                ULONG
locMethodName             BSTRING
locdispid                 LONG
loccArgs                  UNSIGNED
locdispidNamed            LONG(DISPID_PROPERTYPUT)
locpvargtype              GROUP,DIM(1)
pvarg                       VARIANT
gpvarg                      LIKE(gVariant),OVER(pvarg)
                          END
LocRetVal                 HRESULT
HR                        HRESULT

nValue                    LONG
bValue                    BSTRING

  CODE
  
  ASSERT(NOT SELF.m_Disp &= NULL)
  IF SELF.m_Disp &= NULL 
    RETURN
  END
  
  locMethodName=pProp
  HR=SELF.m_Disp.GetIDsOfNames(ADDRESS(IID_NULL),ADDRESS(locMethodName),1,0,ADDRESS(locdispid))
  IF HR <> S_OK
    MESSAGE('"'& pProp &'" property not found, error '& HR)
    RETURN
  END

  locDispParams.cNamedArgs=1
  locDispParams.rgdispidNamedArgs=ADDRESS(locdispidNamed)
  locDispParams.rgvarg=ADDRESS(locpvargtype)
  loccargs+=1
  
  IF NUMERIC(pValue) = TRUE
    nValue = pValue
    locpvargtype[loccargs].pvarg = nValue
    locpvargtype[loccargs].gpvarg.vt = VT_I4
  ELSE
    bValue = pValue
    locpvargtype[loccargs].pvarg = bValue
    locpvargtype[loccargs].gpvarg.vt = VT_BSTR
  END
  
  locDispParams.cArgs=loccargs
  HR=SELF.m_Disp.Invoke(locdispid,ADDRESS(IID_NULL),0,DISPATCH_METHOD+DISPATCH_PROPERTYPUT,ADDRESS(locDispParams),ADDRESS(locgVarResult),locExcepInfo,locuArgErr)
  IF HR < S_OK
    MESSAGE('Get'& pProp &' throws error '& HR)
    RETURN
  END
  
OleItemClass.SetValue   PROCEDURE(STRING pValue)
  CODE
  SELF.SetValue('Value', pValue)
  
!============ OleSafeArrayClass =======================
OleSafeArrayClass.Construct     PROCEDURE()
  CODE

OleSafeArrayClass.Destruct      PROCEDURE()
  CODE
  SELF.Kill()

OleSafeArrayClass.Init          PROCEDURE(*VARIANT pVar)  !,HRESULT,PROC,VIRTUAL
gvArray                         LIKE(gVariant)
vArray                          VARIANT,OVER(gvArray)
  CODE
  SELF.created=FALSE

  vArray=pVar
  SELF.vt = gvArray.vt
  
  IF ~BAND(gvArray.vt,VT_ARRAY) !this is not SAFEARRAY
    RETURN E_INVALIDARG
  END
  SELF.lpSA=gvArray.lval
  RETURN S_OK

OleSafeArrayClass.Init          PROCEDURE(VARTYPE pVTType,LONG pElements) !,HRESULT,PROC,VIRTUAL
rgsabound                       LIKE(_SAFEARRAYBOUND)
  CODE
  SELF.created=TRUE

  SELF.vt = pVTType

  rgsabound.lLbound=1
  rgsabound.cElements=pElements
  SELF.lpSA=ecg_SafeArrayCreate(pVTType,1,ADDRESS(rgsabound))
  IF ~SELF.lpSA
    RETURN E_POINTER
  END
  RETURN S_OK

OleSafeArrayClass.Init          PROCEDURE(VARTYPE pVTType,LONG nDims,LONG[] pElements) !,HRESULT,PROC,VIRTUAL
cDim                            UNSIGNED
rgsabound                       &_SAFEARRAYBOUND
buf                             &STRING
psabound                        LONG
loopIndex                       LONG,AUTO
  CODE
  SELF.created=TRUE

  SELF.vt = pVTType

  cDim=nDims
  buf &= NEW(STRING(cDim*SIZE(_SAFEARRAYBOUND)))  !buffer to store the vector of bounds
  LOOP loopIndex=1 TO cDim
    psabound=ADDRESS(buf)+(loopIndex-1)*SIZE(_SAFEARRAYBOUND)
    rgsabound &= (psabound)
    rgsabound.lLbound=1
    rgsabound.cElements=pElements[loopIndex]
  END
  SELF.lpSA=ecg_SafeArrayCreate(pVTType,cDim,ADDRESS(buf))
  DISPOSE(buf)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END
  RETURN S_OK

OleSafeArrayClass.Kill          PROCEDURE()               !,VIRTUAL
  CODE
  IF SELF.created=TRUE AND SELF.lpSA
    ecg_SafeArrayDestroy(SELF.lpSA)
  END
  SELF.lpSA=0

!OleSafeArrayClass.Destroy       PROCEDURE()               !,VIRTUAL
!  CODE
!  IF SELF.lpSA
!    ecg_SafeArrayDestroy(SELF.lpSA)
!    SELF.lpSA=0
!  END

OleSafeArrayClass.GetVartype    PROCEDURE()               !,USHORT
vt                              VARTYPE
HR                              HRESULT
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN VT_EMPTY
  END

  HR=ecg_SafeArrayGetVartype(SELF.lpSA,vt)
  IF HR <> S_OK
    RETURN VT_EMPTY
  END
  RETURN vt

OleSafeArrayClass.GetElemsize   PROCEDURE() !,UNSIGNED,VIRTUAL
cbSize                            UNSIGNED, AUTO

  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN 0
  END
  cbSize = ecg_SafeArrayGetElemsize(SELF.lpSA)
  RETURN cbSize

OleSafeArrayClass.GetNDims      PROCEDURE()               !,LONG
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN 0
  END
  RETURN ecg_SafeArrayGetDim(SELF.lpSA)

OleSafeArrayClass.GetLBound     PROCEDURE(UNSIGNED pDim)  !,LONG,VIRTUAL
lBound                          LONG
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN 0
  END
  IF ecg_SafeArrayGetLBound(SELF.lpSA,pDim,lBound) <> S_OK
    lBound=0
  END
  RETURN lBound

OleSafeArrayClass.GetUBound     PROCEDURE(UNSIGNED pDim)  !,LONG,VIRTUAL
UBound                          LONG
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN 0
  END
  IF ecg_SafeArrayGetUBound(SELF.lpSA,pDim,UBound) <> S_OK
    UBound=0
  END
  RETURN UBound

OleSafeArrayClass.GetNElements  PROCEDURE(UNSIGNED pDim)      !,LONG
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN 0
  END
  RETURN SELF.GetUBound(pDim)-SELF.GetLBound(pDim)+1

OleSafeArrayClass.GetArray      PROCEDURE() !,LONG,VIRTUAL
  CODE
  RETURN SELF.lpSA

OleSafeArrayClass.GetVariant   PROCEDURE(*VARIANT pv) !,VIRTUAL
gvArray                        LIKE(gVariant)
vArray                         VARIANT,OVER(gvArray)
  CODE
  gvArray.vt=BOR(SELF.GetVartype(),VT_ARRAY)
  gvArray.lVal=SELF.lpSA
  ecg_VariantCopyInd(pv,vArray)

!OleSafeArrayClass.GetElement    PROCEDURE(LONG pDim,LONG pIndex,*VARIANT pElement)  !,HRESULT,VIRTUAL
!rgIndices                       LONG,DIM(2)
!gvElement                       LIKE(gVariant)
!vElement                        VARIANT,OVER(gvElement)
!buf                             STRING(SELF.GetElemsize())
!pData                           LONG,OVER(buf)
!HR                              HRESULT
!  CODE
!  ASSERT(SELF.lpSA)
!  IF ~SELF.lpSA
!    RETURN E_POINTER
!  END
!
!  IF pDim < 1 OR pDim > SELF.GetNDims()
!    RETURN S_FALSE
!  END
!  IF pIndex < 1 OR pIndex > SELF.GetNElements(pDim)
!    RETURN S_FALSE
!  END
!  rgIndices[1]=pIndex+(SELF.GetLBound(pDim)-1)
!  rgIndices[2]=pDim-1
!  HR=ecg_SafeArrayGetElement(SELF.lpSA,ADDRESS(rgIndices),ADDRESS(buf))
!  IF HR <> S_OK
!    RETURN HR
!  END
!  CASE SELF.GetVartype()
!  OF VT_VARIANT
!    gvElement=buf
!  ELSE
!    gvElement.vt=SELF.GetVartype()
!    gvElement.lval=pData
!  END
!  RETURN ecg_VariantCopyInd(pElement,vElement)
OleSafeArrayClass.GetElement    PROCEDURE(LONG pDim,LONG pIndex,*VARIANT pElement)  !,HRESULT,VIRTUAL
rgIndices                         LONG,DIM(2)
gvElement                         LIKE(gVariant)
vElement                          VARIANT,OVER(gvElement)
buf                               &STRING
pData                             LONG
vt                                VARTYPE
HR                                HRESULT
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END

  IF pDim < 1 OR pDim > SELF.GetNDims()
    RETURN S_FALSE
  END
  IF pIndex < 1 OR pIndex > SELF.GetNElements(pDim)
    RETURN S_FALSE
  END
  rgIndices[1]=pIndex+(SELF.GetLBound(pDim)-1)
  rgIndices[2]=pDim-1
  
  vt = SELF.GetVartype()
  CASE vt
  OF VT_VARIANT
    buf &= NEW STRING(SELF.GetElemsize())
    HR=ecg_SafeArrayGetElement(SELF.lpSA, ADDRESS(rgIndices), ADDRESS(buf))
    IF HR = S_OK
      gvElement=buf
    END
    DISPOSE(buf)
    IF HR <> S_OK
      RETURN HR
    END
  ELSE
    HR=ecg_SafeArrayGetElement(SELF.lpSA, ADDRESS(rgIndices), ADDRESS(pData))
    IF HR <> S_OK
      RETURN HR
    END
    gvElement.vt=vt
    gvElement.lval=pData
  END
  RETURN ecg_VariantCopyInd(pElement, vElement)

OleSafeArrayClass.PutElement    PROCEDURE(LONG pDim,LONG pIndex,LONG pElement)  !,HRESULT,VIRTUAL
rgIndices                       LONG,DIM(2)
pData                          &LONG
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END

  IF pDim < 1 OR pDim > SELF.GetNDims()
    RETURN S_FALSE
  END
  IF pIndex < 1 OR pIndex > SELF.GetNElements(pDim)
    RETURN S_FALSE
  END
  rgIndices[1]=pIndex+(SELF.GetLBound(pDim)-1)
  rgIndices[2]=pDim-1
  CASE SELF.GetVartype()
  OF VT_BSTR
    pData &= (pElement)
    RETURN ecg_SafeArrayPutElement(SELF.lpSA,ADDRESS(rgIndices),pData)
  ELSE
    RETURN ecg_SafeArrayPutElement(SELF.lpSA,ADDRESS(rgIndices),pElement)
  END
  
!OleSafeArrayExClass.Init      PROCEDURE(VARTYPE pVTType,LONG pElements,LONG lBound)
!rgsabound                     LIKE(_SAFEARRAYBOUND)
!  CODE
!  SELF.created=TRUE
!
!  rgsabound.lLbound=lBound
!  rgsabound.cElements=pElements
!  SELF.lpSA=ecg_SafeArrayCreate(pVTType,1,ADDRESS(rgsabound))
!  ASSERT(SELF.lpSA)
!  IF ~SELF.lpSA
!    RETURN S_FALSE
!  END
!  RETURN S_OK
OleSafeArrayExClass.Init    PROCEDURE(VARTYPE pVTType, LONG pElements, LONG lBound)
  CODE
  RETURN SELF.Init(pVTType, 1, pElements, lBound)
  
!OleSafeArrayExClass.Init    PROCEDURE(VARTYPE pVTType,LONG pDims, LONG pElements, LONG pLBound)
!rgsabound                     LIKE(_SAFEARRAYBOUND)
!  CODE
!  SELF.created=TRUE
!
!  rgsabound.lLbound=pLBound
!  rgsabound.cElements=pElements
!  SELF.lpSA=ecg_SafeArrayCreate(pVTType,pDims,ADDRESS(rgsabound))
!  ASSERT(SELF.lpSA)
!  IF ~SELF.lpSA
!    RETURN S_FALSE
!  END
!  RETURN S_OK

OleSafeArrayExClass.Init    PROCEDURE(VARTYPE pVTType, LONG pDims, LONG pElements, LONG pLBound)
rgsabound                     &STRING      !-- Pointer to a vector of bounds (one for each dimension) to allocate for the array
sabound                       LIKE(_SAFEARRAYBOUND)
aIndex                        LONG, AUTO
sPos                          LONG, AUTO
ePos                          LONG, AUTO

  CODE
  SELF.created=TRUE
  SELF.vt = pVTType
  
  rgsabound &= NEW(STRING(SIZE(_SAFEARRAYBOUND) * pDims))
  LOOP aIndex = 1 TO pDims
    sPos = (aIndex - 1) * SIZE(_SAFEARRAYBOUND) + 1
    ePos = aIndex * SIZE(_SAFEARRAYBOUND)
    
    sabound = rgsabound[sPos : ePos]
    sabound.lLbound=pLBound               !-- every dim has pLBound lower bound
    sabound.cElements=pElements           !-- every dim has pElements items
    rgsabound[sPos : ePos] = sabound
  END
  
  SELF.lpSA=ecg_SafeArrayCreate(pVTType,pDims,ADDRESS(rgsabound))
  DISPOSE(rgsabound)

  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN S_FALSE
  END
  
  RETURN S_OK

OleSafeArrayExClass.InitQ   PROCEDURE(QUEUE q, LONG pLBound = 0)
cDims                         LONG, AUTO
cElements                     LONG, AUTO
lBound                        LONG, AUTO
a                             ANY
fIndex                        LONG, AUTO
qIndex                        LONG, AUTO
v                             VARIANT
hr                            HRESULT, AUTO

  CODE
  
  !-- calculate number of queue fields
  fIndex = 0
  LOOP
    fIndex += 1
    a &= WHAT(q, fIndex)
    IF a &= NULL
      BREAK
    END
  END
  
  cDims = fIndex - 1
  lBound = pLBound

  ecg::DebugInfo('OleSafeArrayExClass.InitQ: dims '& cDims &', elements '& RECORDS(q))
  
  hr = SELF.Init(VT_VARIANT, cDims, RECORDS(q), lBound)
  ASSERT(hr = S_OK)
  IF hr <> S_OK
    MESSAGE(hr)
    RETURN hr
  END
  
  !-- copy queue into sa
  LOOP qIndex = 1 TO RECORDS(q)
    GET(q, qIndex)
    
    fIndex = 0
    LOOP
      fIndex += 1
      a &= WHAT(q, fIndex)
      IF a &= NULL
        BREAK
      END

      v = a
      hr = SELF.PutVariant(fIndex, qIndex, v)
      IF hr = S_OK
        ecg::DebugInfo('OleSafeArrayExClass.InitQ: PutVariant['& fIndex &', '& qIndex &']='& CLIP(a))
      ELSE
        ecg::DebugInfo('OleSafeArrayExClass.InitQ: PutVariant['& fIndex &', '& qIndex &'] failed, hr = '& hr)
      END
    END
  END

  RETURN S_OK

OleSafeArrayExClass.InitQ1  PROCEDURE(QUEUE q, LONG pLBound = 0)
nFields                       LONG, AUTO
cElements                     LONG, AUTO
lBound                        LONG, AUTO
a                             ANY
fIndex                        LONG, AUTO
qIndex                        LONG, AUTO
v                             VARIANT
hr                            HRESULT, AUTO

  CODE
  
  !-- calculate number of queue fields
  fIndex = 0
  LOOP
    fIndex += 1
    a &= WHAT(q, fIndex)
    IF a &= NULL
      BREAK
    END
  END
  
  nFields = fIndex - 1
  lBound = pLBound

  hr = SELF.Init(VT_VARIANT, 1, RECORDS(q) * nFields, lBound)
  ASSERT(hr = S_OK)
  IF hr <> S_OK
    MESSAGE(hr)
    RETURN hr
  END
  
  !-- copy queue into sa
  LOOP qIndex = 1 TO RECORDS(q)
    GET(q, qIndex)
    
    fIndex = 0
    LOOP
      fIndex += 1
      a &= WHAT(q, fIndex)
      IF a &= NULL
        BREAK
      END

      v = a
      SELF.PutVariant(1, nFields * (qIndex - 1) + fIndex, v)
    END
  END

  RETURN S_OK

OleSafeArrayExClass.InitAsPPSA  PROCEDURE(LONG ppSA)  ! ppSA = SAFEARRAY**
  CODE
  SELF.created = FALSE
  ecg_memcpy(ADDRESS(SELF.lpSA), ppSA, 4)
  RETURN S_OK
  
OleSafeArrayExClass.PutVariant    PROCEDURE(LONG pDim,LONG pIndex,*VARIANT pVar, LONG vt=VT_EMPTY)
rgIndices                         LONG,DIM(2)
v                                 VARIANT
gv                                LIKE(gVariant),OVER(v)
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END

  ASSERT(SELF.GetVartype()=VT_VARIANT)
  IF SELF.GetVartype() <> VT_VARIANT
    RETURN S_FALSE
  END

  IF pDim < 1 OR pDim > SELF.GetNDims()
    RETURN S_FALSE
  END
  IF pIndex < 1 OR pIndex > SELF.GetNElements(pDim)
    RETURN S_FALSE
  END

  rgIndices[1]=pIndex+(SELF.GetLBound(pDim)-1)
  rgIndices[2]=pDim-1
  
  ecg::DebugInfo('OleSafeArrayExClass.PutVariant: rgIndices[1] '& rgIndices[1] &', rgIndices[2] '& rgIndices[2])
  
  IF vt <> VT_BSTR
    v=pVar
  ELSE
    v=CLIP(pVar)
  END
  IF vt <> VT_EMPTY
    gv.vt=vt
  END

  ecg_VariantClear(pVar)

  RETURN ecg_SafeArrayPutElement(SELF.lpSA,ADDRESS(rgIndices),ADDRESS(gv))
  
OleSafeArrayExClass.ReDim       PROCEDURE(LONG newSize)
rgsabound                       LIKE(_SAFEARRAYBOUND)  
  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END
  ASSERT(SELF.GetNDims() = 1)
  IF ~(SELF.GetNDims() = 1)
    RETURN S_FALSE
  END
  IF SELF.GetNElements(1) = newSize
    RETURN S_OK
  END
  rgsabound.lLbound=SELF.GetLBound(1)
  rgsabound.cElements=newSize
  RETURN ecg_SafeArrayRedim(SELF.lpSA, ADDRESS(rgsabound))
  
OleSafeArrayExClass.CopyToVariant   PROCEDURE()
gv                                  LIKE(gVariant)
v                                   VARIANT,OVER(gv)
  CODE
  ASSERT(SELF.lpSA)
!  STOP(SELF.vt &' '& SELF.lpSA)
  IF SELF.lpSA
    gv.lval=SELF.lpSA
    gv.vt=BOR(VT_ARRAY,SELF.vt)
  END
  RETURN v

OleSafeArrayExClass.CopyToVariant   PROCEDURE(*gVariant gv)
  CODE
  ASSERT(SELF.lpSA)
  ASSERT(SELF.vt > VT_NULL)
  IF SELF.lpSA
    gv.lval=SELF.lpSA
    gv.vt=BOR(VT_ARRAY,SELF.vt)
  END

OleSafeArrayExClass.Destroy         PROCEDURE(LONG psa)
  CODE
  IF psa
    RETURN ecg_SafeArrayDestroy(psa)
  END
  RETURN S_OK

OleSafeArrayExClass.GetData PROCEDURE(LONG pBuffer, LONG pBufferSize)   !,HRESULT,PROC
pvData                        LONG, AUTO
HR                            HRESULT

  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END
  
  HR = ecg_SafeArrayAccessData(SELF.lpSA, pvData)
  IF HR = S_OK
    ecg_memcpy(pBuffer, pvData, pBufferSize)
    ecg_SafeArrayUnaccessData(SELF.lpSA)
  END
  
  RETURN HR
  
OleSafeArrayExClass.PutData PROCEDURE(LONG pBuffer, LONG pBufferSize)   !,HRESULT,PROC
pvData                        LONG, AUTO
HR                            HRESULT

  CODE
  ASSERT(SELF.lpSA)
  IF ~SELF.lpSA
    RETURN E_POINTER
  END
  
  HR = ecg_SafeArrayAccessData(SELF.lpSA, pvData)
  IF HR = S_OK
    ecg_memcpy(pvData, pBuffer, pBufferSize)
    ecg_SafeArrayUnaccessData(SELF.lpSA)
  END
  
  RETURN HR

!============ OleVariantClass =======================
OleVariantClass.GetValue    PROCEDURE(? par)
v                           VARIANT
hr                          HRESULT
  CODE
  v=par   !vt=VT_BSTR
  IF NOT ISSTRING(par)
    hr=ecg_VariantChangeType(v, v, , GetVT_Type(par))
    IF hr <> S_OK
      STOP('vt(' & par & ')=' & hr)
    END
  END
  RETURN v

OleVariantClass.Missing     PROCEDURE()
v                           VARIANT
gv                          LIKE(gVariant),OVER(v)
  CODE
  !gv.vt=VT_ERROR
  gv.vt=VT_EMPTY
  RETURN v

OleVariantClass.GetAddress  PROCEDURE(? par)
v                           VARIANT
gv                          LIKE(gVariant),OVER(v)

g                           &STRING
  CODE
  v=par
  g &= NEW(STRING(SIZE(gv)))
  g = gv
  RETURN ADDRESS(g)

!============ OleWrapperExClass =======================

OleWrapperExClass.Construct PROCEDURE()
  CODE
  SELF.IsLibInitialized=CHOOSE(CoInitialize() >= 0, TRUE, FALSE)

OleWrapperExClass.Init  PROCEDURE(STRING ProgId,STRING ClassName,SIGNED Ctrl)
  CODE
  SELF.ClassName=CLIP(ClassName)
  SELF.debug=1

  IF NOT SELF.IsLibInitialized
    IF SELF.debug
      SELF.ShowError('COM initializing',S_FALSE)
    END
    RETURN 0
  END

  SELF.OleCtrl=CREATE(0,CREATE:OLE)
  SELF.OleCtrl{prop:create}=ProgId
  SELF.ParentCtrl=Ctrl
  SELF.Resize();
  UNHIDE(SELF.OleCtrl)
  RETURN SELF.OleCtrl
