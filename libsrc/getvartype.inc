  SECTION('DECLARATIONS')
    GetVT_Type(*BYTE),LONG
    GetVT_Type(*SHORT),LONG
    GetVT_Type(*USHORT),LONG
    GetVT_Type(*DATE),LONG
    GetVT_Type(*TIME),LONG
    GetVT_Type(*LONG),LONG
    GetVT_Type(*ULONG),LONG
    GetVT_Type(*SREAL),LONG
    GetVT_Type(*REAL),LONG 
    GetVT_Type(*BFLOAT4),LONG 
    GetVT_Type(*BFLOAT8),LONG 
    GetVT_Type(*BSTRING),LONG 
    GetVT_Type(*_SAFEARRAY),LONG 
    GetVT_Type(*VARIANT),LONG 
    GetVT_Type(*?),LONG
  SECTION('CODE')
GetVT_Type                      PROCEDURE(*BYTE pVar)
  CODE
  RETURN VT_UI1
GetVT_Type                      PROCEDURE(*SHORT pVar)
  CODE
  RETURN VT_I2
GetVT_Type                      PROCEDURE(*USHORT pVar)
  CODE
  RETURN VT_UI2
GetVT_Type                      PROCEDURE(*DATE pVar)
  CODE
  RETURN VT_BSTR
GetVT_Type                      PROCEDURE(*TIME pVar)
  CODE
  RETURN VT_BSTR
GetVT_Type                      PROCEDURE(*LONG pVar)
  CODE
  RETURN VT_I4
GetVT_Type                      PROCEDURE(*ULONG pVar)
  CODE
  RETURN VT_UI4
GetVT_Type                      PROCEDURE(*SREAL pVar)
  CODE
  RETURN VT_R4
GetVT_Type                      PROCEDURE(*REAL pVar)
  CODE
  RETURN VT_R8
GetVT_Type                      PROCEDURE(*BFLOAT4 pVar)
  CODE
  RETURN VT_R4
GetVT_Type                      PROCEDURE(*BFLOAT8 pVar)
  CODE
  RETURN VT_R8
GetVT_Type                      PROCEDURE(*BSTRING pVar)
  CODE
  RETURN VT_BSTR
GetVT_Type                      PROCEDURE(*_SAFEARRAY pVar)
  CODE
  RETURN VT_ARRAY
  !RETURN VT_BSTR
GetVT_Type                      PROCEDURE(*VARIANT pVar)
v                               VARIANT
gv                              LIKE(gVariant),OVER(v)
  CODE
  v=pVar
  RETURN gv.vt
GetVT_Type                      PROCEDURE(*? pVar)
  CODE
  RETURN VT_EMPTY
