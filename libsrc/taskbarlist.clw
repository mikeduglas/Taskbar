!Generated .CLW file (by the Easy COM generator v 1.14)

  MEMBER
  INCLUDE('svcomdef.inc'),ONCE
  MAP
    MODULE('WinAPI')
      ecg_DispGetParam(LONG pdispparams,LONG dwPosition,VARTYPE vtTarg,*VARIANT pvarResult,*SIGNED uArgErr),HRESULT,RAW,PASCAL,NAME('DispGetParam')
      ecg_VariantInit(*VARIANT pvarg),HRESULT,PASCAL,PROC,NAME('VariantInit')
      ecg_VariantClear(*VARIANT pvarg),HRESULT,PASCAL,PROC,NAME('VariantClear')
      ecg_VariantCopy(*VARIANT vargDest,*VARIANT vargSrc),HRESULT,PASCAL,PROC,NAME('VariantCopy')
      memcpy(LONG lpDest,LONG lpSource,LONG nCount),LONG,PROC,NAME('_memcpy')
      GetErrorInfo(ULONG dwReserved,LONG pperrinfo),HRESULT,PASCAL,PROC
    END
    INCLUDE('svapifnc.inc')
    Dec2Hex(ULONG),STRING
    INCLUDE('getvartype.inc', 'DECLARATIONS')
  END
  INCLUDE('taskbarlist.inc')

Dec2Hex                       PROCEDURE(ULONG pDec)
locHex                          STRING(30)
  CODE
  LOOP UNTIL(~pDec)
    locHex = SUB('0123456789ABCDEF',1+pDec % 16,1) & CLIP(locHex)
    pDec = INT(pDec / 16)
  END
  RETURN CLIP(locHex)

  INCLUDE('getvartype.inc', 'CODE')
!========================================================!
! ITaskbarList3Class implementation                      !
!========================================================!
ITaskbarList3Class.Construct  PROCEDURE()
  CODE
  SELF.debug=true

ITaskbarList3Class.Destruct   PROCEDURE()
  CODE
  IF SELF.IsInitialized=true THEN SELF.Kill() END

ITaskbarList3Class.Init       PROCEDURE()
loc:lpInterface                 LONG
  CODE
  SELF.HR=CoCreateInstance(ADDRESS(IID_TaskbarInstance),0,SELF.dwClsContext,ADDRESS(IID_ITaskbarList3),loc:lpInterface)
  IF SELF.HR=S_OK
    RETURN SELF.Init(loc:lpInterface)
  ELSE
    SELF.IsInitialized=false
    SELF._ShowErrorMessage('ITaskbarList3Class.Init: CoCreateInstance',SELF.HR)
  END
  RETURN SELF.HR

ITaskbarList3Class.Init       PROCEDURE(LONG lpInterface)
  CODE
  IF PARENT.Init(lpInterface) = S_OK
    SELF.ITaskbarList3Obj &= (lpInterface)
  END
  RETURN SELF.HR

ITaskbarList3Class.Kill       PROCEDURE()
  CODE
  IF PARENT.Kill() = S_OK
    SELF.ITaskbarList3Obj &= NULL
  END
  RETURN SELF.HR

ITaskbarList3Class.GetInterfaceObject PROCEDURE()
  CODE
  RETURN SELF.ITaskbarList3Obj

ITaskbarList3Class.GetInterfaceAddr   PROCEDURE()
  CODE
  RETURN ADDRESS(SELF.ITaskbarList3Obj)
  !RETURN INSTANCE(SELF.ITaskbarList3Obj, 0)

ITaskbarList3Class.GetLibLocation PROCEDURE()
  CODE
  RETURN GETREG(REG_CLASSES_ROOT,'CLSID\{{56fdf344-fd6d-11d0-958a-006097c9a090}\InprocServer32')

ITaskbarList3Class.HrInit     PROCEDURE()
HR                              HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.HrInit()
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.HrInit',HR)
  END
  RETURN HR

ITaskbarList3Class.AddTab     PROCEDURE(HWND phWnd)
HR                              HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.AddTab(phWnd)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.AddTab',HR)
  END
  RETURN HR

ITaskbarList3Class.DeleteTab  PROCEDURE(HWND phWnd)
HR                              HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.DeleteTab(phWnd)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.DeleteTab',HR)
  END
  RETURN HR

ITaskbarList3Class.ActivateTab    PROCEDURE(HWND phWnd)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.ActivateTab(phWnd)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.ActivateTab',HR)
  END
  RETURN HR

ITaskbarList3Class.SetActiveAlt   PROCEDURE(HWND phWnd)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetActiveAlt(phWnd)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetActiveAlt',HR)
  END
  RETURN HR

ITaskbarList3Class.MarkFullscreenWindow   PROCEDURE(HWND phWnd,bool pfFullscreen)
HR                                          HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.MarkFullscreenWindow(phWnd,pfFullscreen)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.MarkFullscreenWindow',HR)
  END
  RETURN HR

ITaskbarList3Class.SetProgressValue   PROCEDURE(HWND phWnd,long ullCompletedLo,long ullCompletedHi,long ullTotalLo,long ullTotalHi)
HR                                      HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetProgressValue(phWnd,ullCompletedLo,ullCompletedHi,ullTotalLo,ullTotalHi)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetProgressValue',HR)
  END
  RETURN HR

ITaskbarList3Class.SetProgressState   PROCEDURE(HWND phWnd,TBPFLAG ptbpFlags)
HR                                      HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetProgressState(phWnd,ptbpFlags)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetProgressState',HR)
  END
  RETURN HR

ITaskbarList3Class.RegisterTab    PROCEDURE(HWND phWndTab,HWND phWndMDI)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.RegisterTab(phWndTab,phWndMDI)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.RegisterTab',HR)
  END
  RETURN HR

ITaskbarList3Class.UnregisterTab  PROCEDURE(HWND phWndTab)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.UnregisterTab(phWndTab)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.UnregisterTab',HR)
  END
  RETURN HR

ITaskbarList3Class.SetTabOrder    PROCEDURE(HWND phWndTab,HWND phWndInsertBefore)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetTabOrder(phWndTab,phWndInsertBefore)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetTabOrder',HR)
  END
  RETURN HR

ITaskbarList3Class.SetTabActive   PROCEDURE(HWND phWndTab,HWND phWndMDI,ulong ptbatFlags)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetTabActive(phWndTab,phWndMDI,ptbatFlags)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetTabActive',HR)
  END
  RETURN HR

ITaskbarList3Class.ThumbBarAddButtons PROCEDURE(HWND phWnd,ulong pcButtons,long ppButton)
HR                                      HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.ThumbBarAddButtons(phWnd,pcButtons,ppButton)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.ThumbBarAddButtons',HR)
  END
  RETURN HR

ITaskbarList3Class.ThumbBarUpdateButtons  PROCEDURE(HWND phWnd,ulong pcButtons,long ppButton)
HR                                          HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.ThumbBarUpdateButtons(phWnd,pcButtons,ppButton)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.ThumbBarUpdateButtons',HR)
  END
  RETURN HR

ITaskbarList3Class.ThumbBarSetImageList   PROCEDURE(HWND phWnd,HANDLE phiml)
HR                                          HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.ThumbBarSetImageList(phWnd,phiml)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.ThumbBarSetImageList',HR)
  END
  RETURN HR

ITaskbarList3Class.SetOverlayIcon PROCEDURE(HWND phWnd,HANDLE phIcon,BSTRING ppszDescription)
HR                                  HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetOverlayIcon(phWnd,phIcon,ppszDescription)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetOverlayIcon',HR)
  END
  RETURN HR

ITaskbarList3Class.SetThumbnailTooltip    PROCEDURE(HWND phWnd,BSTRING ppszTip)
HR                                          HRESULT(S_OK)

  CODE
  IF SELF.IsInitialized=false THEN SELF.HR=E_FAIL;RETURN(SELF.HR) END
  HR=SELF.ITaskbarList3Obj.SetThumbnailTooltip(phWnd,ppszTip)
  SELF.HR=HR
  IF HR < S_OK
    SELF._ShowErrorMessage('ITaskbarList3Class.SetThumbnailTooltip',HR)
  END
  RETURN HR


!region TTaskbarProgress
TTaskbarProgress.Construct    PROCEDURE()
  CODE
  SELF.m_COMIniter &= NEW CCOMIniter
  SELF.debug = FALSE
  SELF.Init()

TTaskbarProgress.Destruct    PROCEDURE()
  CODE
  PARENT.Destruct()
  DISPOSE(SELF.m_COMIniter)

TTaskbarProgress.IsSupported  PROCEDURE()
  CODE
  RETURN SELF.IsInitialized
  
TTaskbarProgress.SetWindow    PROCEDURE(WINDOW pWin)
  CODE
  SELF.hwnd = pWin{PROP:Handle}
                                
TTaskbarProgress.SetValue     PROCEDURE(LONG ullCompleted, LONG ullTotal)
hwnd                            HWND, AUTO
  CODE
  hwnd = CHOOSE(SELF.hwnd <> 0, SELF.hwnd, 0{PROP:Handle})
  RETURN CHOOSE(PARENT.SetProgressValue(hwnd, ullCompleted, 0, ullTotal, 0))

TTaskbarProgress.SetPercent   PROCEDURE(BYTE pCompleted)
  CODE
  RETURN SELF.SetValue(pCompleted, 100)
  
TTaskbarProgress.SetState     PROCEDURE(TBPFLAG tbpFlags)
hwnd                            HWND, AUTO
  CODE
  hwnd = CHOOSE(SELF.hwnd <> 0, SELF.hwnd, 0{PROP:Handle})
  RETURN CHOOSE(PARENT.SetProgressState(hwnd, tbpFlags))
!endregion