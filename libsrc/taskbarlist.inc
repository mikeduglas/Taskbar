!Generated .INC file (by the Easy COM generator v 1.14)

  INCLUDE('ecom2inc.def'),ONCE
  INCLUDE('taskbarlist.int'),ONCE
  INCLUDE('ecombase.inc'),ONCE

ITaskbarList3Class            CLASS(EasyCOMClass),TYPE,MODULE('taskbarlist.clw'),LINK('taskbarlist.clw')
ITaskbarList3Obj                &ITaskbarList3,PROTECTED
Construct                       PROCEDURE()
Destruct                        PROCEDURE(),DERIVED
Init                            PROCEDURE(),HRESULT,PROC,DERIVED
Init                            PROCEDURE(LONG lpInterface),HRESULT,PROC,DERIVED
Kill                            PROCEDURE(),HRESULT,PROC,DERIVED
GetInterfaceObject              PROCEDURE(),*ITaskbarList3
GetInterfaceAddr                PROCEDURE(),LONG
GetLibLocation                  PROCEDURE(),STRING,DERIVED
!Interface methods
!!! <summary>
!!! Initializes the taskbar list object. This method must be called before any other ITaskbarList methods can be called.
!!! </summary>
HrInit                          PROCEDURE(),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Adds an item to the taskbar.
!!! </summary>
AddTab                          PROCEDURE(HWND hWnd),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Deletes an item from the taskbar.
!!! </summary>
DeleteTab                       PROCEDURE(HWND hWnd),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Activates an item on the taskbar. The window is not actually activated; the window's item on the taskbar is merely displayed as active.
!!! </summary>
ActivateTab                     PROCEDURE(HWND hWnd),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Marks a taskbar item as active but does not visually activate it.
!!! </summary>
SetActiveAlt                    PROCEDURE(HWND hWnd),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Marks a window as full-screen.
!!! </summary>
MarkFullscreenWindow            PROCEDURE(HWND hWnd,bool fFullscreen),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Displays or updates a progress bar hosted in a taskbar button to show the specific percentage completed of the full operation.
!!! </summary>
SetProgressValue                PROCEDURE(HWND hWnd,long ullCompletedLo,long ullCompletedHi,long ullTotalLo,long ullTotalHi),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Sets the type and state of the progress indicator displayed on a taskbar button.
!!! </summary>
SetProgressState                PROCEDURE(HWND hWnd,TBPFLAG tbpFlags),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Informs the taskbar that a new tab or document thumbnail has been provided for display in an application's taskbar group flyout.
!!! </summary>
RegisterTab                     PROCEDURE(HWND hWndTab,HWND hWndMDI),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Removes a thumbnail from an application's preview group when that tab or document is closed in the application.
!!! </summary>
UnregisterTab                   PROCEDURE(HWND hWndTab),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Inserts a new thumbnail into a tabbed-document interface (TDI) or multiple-document interface (MDI) application's group flyout or moves an existing thumbnail to a new position in the application's group.
!!! </summary>
SetTabOrder                     PROCEDURE(HWND hWndTab,HWND hWndInsertBefore),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Informs the taskbar that a tab or document window has been made the active window.
!!! </summary>
SetTabActive                    PROCEDURE(HWND hWndTab,HWND hWndMDI,ulong tbatFlags),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Adds a thumbnail toolbar with a specified set of buttons to the thumbnail image of a window in a taskbar button flyout.
!!! </summary>
ThumbBarAddButtons              PROCEDURE(HWND hWnd,ulong cButtons,long pButton),HRESULT,PROC,VIRTUAL
ThumbBarUpdateButtons           PROCEDURE(HWND hWnd,ulong cButtons,long pButton),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Specifies an image list that contains button images for a toolbar embedded in a thumbnail image of a window in a taskbar button flyout.
!!! </summary>
ThumbBarSetImageList            PROCEDURE(HWND hWnd,HANDLE himl),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Applies an overlay to a taskbar button to indicate application status or a notification to the user.
!!! </summary>
SetOverlayIcon                  PROCEDURE(HWND hWnd,HANDLE hIcon,BSTRING pszDescription),HRESULT,PROC,VIRTUAL
!!! <summary>
!!! Specifies or updates the text of the tooltip that is displayed when the mouse pointer rests on an individual preview thumbnail in a taskbar button flyout.
!!! </summary>
SetThumbnailTooltip             PROCEDURE(HWND hWnd,BSTRING pszTip),HRESULT,PROC,VIRTUAL
                              END


TTaskbarProgress              CLASS(ITaskbarList3Class), TYPE, MODULE('taskbarlist.clw'), LINK('taskbarlist.clw')
m_COMIniter                     &CCOMIniter, PRIVATE
hwnd                            HWND, PRIVATE

Construct                       PROCEDURE()
Destruct                        PROCEDURE(), DERIVED

IsSupported                     PROCEDURE(), BOOL

!!! <summary>
!!! Sets the window in taskbar. If not called, SetValue and SetState will use last opened window.
!!! </summary>
SetWindow                       PROCEDURE(WINDOW pWin)
!!! <summary>
!!! Displays or updates a progress bar hosted in a taskbar button to show the specific percentage completed of the full operation.
!!! </summary>
SetValue                        PROCEDURE(LONG ullCompleted, LONG ullTotal), BOOL, PROC
!!! <summary>
!!! Displays or updates a progress bar hosted in a taskbar button to show the specific percentage completed of the full operation.
!!! </summary>
SetPercent                      PROCEDURE(BYTE pCompleted), BOOL, PROC
!!! <summary>
!!! Sets the type and state of the progress indicator displayed on a taskbar button.
!!! </summary>
SetState                        PROCEDURE(TBPFLAG tbpFlags), BOOL, PROC
                              END
