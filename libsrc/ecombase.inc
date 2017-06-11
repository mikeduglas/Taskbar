  INCLUDE('svcom.inc'),ONCE
  INCLUDE('easyole.inc'),ONCE

!Base class for gen'd com classes
  
EasyCOMClass          CLASS(CCOMObject),TYPE,MODULE('ecombase.clw'),LINK('ecombase.clw')
ClassName               CSTRING(101),PROTECTED  !actual class name, must be set in ctor
dwClsContext            ULONG,PROTECTED         !CoCreateInstance() argument, can be overridden in ctor
CVariant                &OleVariantClass,PROTECTED
IsInitialized           BOOL,PROTECTED
debug                   BOOL
HR                      HRESULT,PROTECTED
Construct               PROCEDURE()
Destruct                PROCEDURE(),VIRTUAL
Init                    PROCEDURE(),HRESULT,PROC,VIRTUAL  !stub
Init                    PROCEDURE(REFCLSID clsid, REFIID iid),HRESULT,PROC,VIRTUAL,PROTECTED  !is called from Init()
Init                    PROCEDURE(LONG lpInterface),HRESULT,PROC,VIRTUAL
Init                    PROCEDURE(*IUnknown pInterface),HRESULT,PROC,VIRTUAL
Kill                    PROCEDURE(),HRESULT,PROC,VIRTUAL
GetLibLocation          PROCEDURE(),STRING,VIRTUAL        !returns registry key CLSID\{...}\InprocServer32
GetNativeError          PROCEDURE(*BSTRING pErrorStr),VIRTUAL,PROTECTED
GetNativeError          PROCEDURE(),STRING
_ShowErrorMessage       PROCEDURE(STRING pMethodName,HRESULT pHR,BYTE pUseDebugMode=TRUE),HRESULT,PROC,VIRTUAL
Dec2Hex                 PROCEDURE(ULONG pDec,BYTE pAsString=0),STRING,PROTECTED
GetDispId               PROCEDURE(STRING pMethodName, long lcid=0, *LONG dispid),HRESULT
                      END

