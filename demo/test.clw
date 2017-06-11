  PROGRAM

  INCLUDE('taskbarlist.inc'), ONCE

  MAP
  END

tbProgress                    TTaskbarProgress

!-- Taskbar states
fState                        TBPFLAG(TBPF_NORMAL)

!-- Progress vars
nTotals                       LONG(100)
nCompleted                    LONG(50)

Window                        WINDOW('Progressbar in Taskbar'),AT(,,305,128),CENTER,GRAY,ICON('image.ico'), |
                                FONT('Segoe UI',10)
                                PROMPT('Progress bar value:'),AT(18,10),USE(?PROMPT1)
                                SLIDER,AT(18,27,273),USE(nCompleted),RANGE(0,100),BELOW
                                OPTION('Taskbar state'),AT(18,53,188,59),USE(fState),BOXED
                                  RADIO('No progress'),AT(36,66),USE(?fState:NOPROGRESS)
                                  RADIO('Indeterminate'),AT(129,66),USE(?fState:INDETERMINATE)
                                  RADIO('Normal'),AT(36,88),USE(?fState:NORMAL)
                                  RADIO('Error'),AT(96,88),USE(?fState:ERROR)
                                  RADIO('Paused'),AT(145,88),USE(?fState:PAUSED)
                                END
                                BUTTON('Exit'),AT(239,98,52),USE(?bExit),STD(STD:Close)
                              END

  CODE
  IF NOT tbProgress.IsSupported()
    MESSAGE('Not supported.', 'Error', ICON:Exclamation)
    RETURN
  END
    
  OPEN(Window)
  
  !-- init controls
  
  !-- radio button true values
  ?fState:NOPROGRESS{PROP:TrueValue} = TBPF_NOPROGRESS
  ?fState:INDETERMINATE{PROP:TrueValue} = TBPF_INDETERMINATE
  ?fState:NORMAL{PROP:TrueValue} = TBPF_NORMAL
  ?fState:ERROR{PROP:TrueValue} = TBPF_ERROR
  ?fState:PAUSED{PROP:TrueValue} = TBPF_PAUSED
  
  !-- slider min/max
  ?nCompleted{PROP:RangeLow} = 0
  ?nCompleted{PROP:RangeHigh} = nTotals
  ?nCompleted{PROP:Step} = 1
  
  !-- set window
  tbProgress.SetWindow(Window)
  
  ACCEPT
    CASE EVENT()
    OF EVENT:OpenWindow
      !-- set progress to some non-zero value -- just for demonstration
      tbProgress.SetValue(nCompleted, nTotals)
      tbProgress.SetState(fState)
      
    OF EVENT:Accepted
      CASE FIELD()
      OF ?fState
        tbProgress.SetState(fState)
     
      OF ?nCompleted
        tbProgress.SetValue(nCompleted, nTotals)
        !tbProgress.SetPercent(nCompleted)
      END
    END
  END
