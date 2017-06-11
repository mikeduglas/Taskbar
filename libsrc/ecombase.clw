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
      ecg_SafeArrayGetElement(LONG psa,LONG prgIndices,LONG pv),LONG,PASCAL,NAME('SafeArrayGetElement')
      ecg_SafeArrayPutElement(LONG psa,LONG prgIndices,LONG pv),LONG,PASCAL,NAME('SafeArrayPutElement')
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

      ecg_lstrlen(LONG lpWideString),LONG,NAME('lstrlenw'),PASCAL   !Randy Rogers' fix

      ecg_CoInitializeEx(LONG,ULONG),HRESULT,RAW,PASCAL,NAME('CoInitializeEx')
      ecg_GetErrorInfo(ULONG dwReserved,LONG pperrinfo),HRESULT,PASCAL,PROC,NAME('GetErrorInfo')
    END
  END
  INCLUDE('ecom2inc.def'),ONCE
  INCLUDE('ecombase.inc'),ONCE

!============ EasyCOMClass =======================
EasyCOMClass.Construct          PROCEDURE()
  CODE
  SELF.CVariant &= NEW (OleVariantClass)

  SELF.debug=true
  SELF.dwClsContext=CLSCTX_ALL_DCOM

EasyCOMClass.Destruct           PROCEDURE()
  CODE
  DISPOSE(SELF.CVariant)

  IF SELF.IsInitialized=true THEN SELF.Kill() END

EasyCOMClass.Init               PROCEDURE()
  CODE
  RETURN E_NOTIMPL

EasyCOMClass.Init               PROCEDURE(REFCLSID clsid, REFIID iid)
lpInterface                     LONG
  CODE
  SELF.HR=CoCreateInstance(clsid,0,SELF.dwClsContext,iid,lpInterface)
  IF SELF.HR=S_OK
    SELF.HR=SELF.Init(lpInterface)
    IF SELF.HR=S_OK
      SELF.IsInitialized=true
    END
  ELSE
    SELF.IsInitialized=false
    SELF._ShowErrorMessage(SELF.ClassName & '.Init: CoCreateInstance',SELF.HR)
  END
  RETURN SELF.HR

EasyCOMClass.Init               PROCEDURE(LONG lpInterface)
  CODE
  ASSERT(lpInterface <> 0)
  IF lpInterface <> 0
    SELF.IUnk &= (lpInterface)
    SELF.IsInitialized=true
    SELF.HR=S_OK
  ELSE
    SELF.HR=E_NOINTERFACE
  END
  RETURN SELF.HR

EasyCOMClass.Init               PROCEDURE(*IUnknown pInterface)
  CODE
  ASSERT(NOT pInterface &= NULL)
  IF NOT pInterface &= NULL
    SELF.IUnk &= pInterface
    SELF.IsInitialized=true
    SELF.HR=S_OK
  ELSE
    SELF.HR=E_NOINTERFACE
  END
  RETURN SELF.HR

EasyCOMClass.Kill               PROCEDURE()
nCnt                            LONG,AUTO
  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_NOINTERFACE;RETURN(SELF.HR) END
!  LOOP
!    nCnt = SELF.Release()
!    IF nCnt=0
!      BREAK
!    END
!  END
  SELF.IUnk &= NULL
  SELF.IsInitialized=false
  SELF.HR=S_OK
  RETURN SELF.HR

EasyCOMClass.GetLibLocation     PROCEDURE()
  CODE
  RETURN ''

EasyCOMClass.GetNativeError     PROCEDURE(*BSTRING pErrorStr)
pError                          &IErrorInfo
lpErrorInfo                     LONG
HR                              HRESULT
  CODE
  HR=ecg_GetErrorInfo(0,ADDRESS(lpErrorInfo))
  IF HR=S_OK AND lpErrorInfo
    pError &= (lpErrorInfo)
    SELF.HR=pError.GetDescription(pErrorStr)
  ELSE
    CLEAR(pErrorStr)
  END

EasyCOMClass.GetNativeError PROCEDURE()
loc:NativeError               BSTRING
  CODE
  SELF.GetNativeError(loc:NativeError)
  RETURN CLIP(loc:NativeError)
  
EasyCOMClass._ShowErrorMessage  PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=TRUE)
  CODE
  IF SELF.debug=TRUE OR pUseDebugMode=FALSE
    MESSAGE(SELF.ClassName & '.' & pMethodName & ' failed (' & SELF.Dec2Hex(pHR) & ')<13,10>' & SELF.GetNativeError(),'COM Error',ICON:EXCLAMATION,,,MSGMODE:CANCOPY)
  END

  RETURN S_OK

EasyCOMClass.Dec2Hex            PROCEDURE(ULONG pDec,BYTE pAsString)
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

EasyCOMClass.GetDispId          PROCEDURE(STRING pMethodName, long lcid=0, *LONG dispid)
bMethodName                     BSTRING
idisp                           &IDispatch
pdisp                           LONG
  CODE
  SELF.HR = SELF.IUnk.QueryInterface(ADDRESS(_IDispatch), pdisp)
  IF SELF.HR = S_OK
    idisp &= (pdisp)
    bMethodName=CLIP(pMethodName)
    SELF.HR = idisp.GetIDsOfNames(ADDRESS(IID_NULL),ADDRESS(bMethodName),1,lcid,ADDRESS(dispid))
  END
  RETURN SELF.HR

