;;Opt("MustDeclareVars", 1)
#AutoIt3Wrapper_Au3Check_Parameters= -w 4 Jos

#cs
   UDF cfx.au3
   serial functions using kernel32.dll
   V1.0
   Uwe Lahni 2008
   V2.0
   Andrew Calcutt 05/16/2009
		 Started converting to UDF
   V2.1
   Mikko Keski-Heroja 02/23/2011
		 UDF is now compatible with Opt("MustDeclareVars",1) and Date.au3.
   V2.2
   Veronesi 04/26/2011
		 Changed some cosmetics and documentation
		 Add Function to set RTS and to get DCD Status
   V2.90
   Maze 07/01/2017
	  Renamed to: ComUDF.au3
	  Deleted functions to set RTS and to get DCD Status
	  Only one global variable: $__ComUDFdll (kernel32.dll)
	  Changed COM port format ("COM" plus number, e.g. "COM1")
	  Changed all function names to start with '_Com'
	  Changed function _ComOpenPort to accept only one argument $defs
		 It gives full flexibility while keeping the amount of parameters small
	  Modified and corrected the parameters for DLLCall
		 Each structure and pointer to that structure is manually defined, according
		 to the names provided in the MSDN reference. Each DllCall uses these names.

		 (While this looks much longer and more complex, it is easier to write new
		  functions in this format and it is also easier to read and check existing
		  functions. It avoids shortcuts therefore can access all data directly from
		  the provided struct. Using shortcuts doesn't allow this and requires access
		  to the correct index of the returned array. (Look for example at the
		  _WinAPI_WriteFile which has a statement that says "$iWritten = $aResult[4]".
		  It takes a while to verify the "4" is indeed the correct index.)
		  There is only one exception regarding the naming, namely the handle to the
		  serial port named "hComPort" and used instead of "hFile".)

	  Added a bunch of new functions, removed obsolete functions
	  Added/modified comments and made a few cosmetic changes
#ce

; #INDEX# =======================================================================================================================
; Title .........: Serial communication UDF
; AutoIt Version : 3.3.14.2
; Description ...: Serial communication using kernel32.dll (no custom dll needed)
; Authors........: Uwe Lahni, Andrew Calcutt, Mikko Keski-Heroja, Veronesi, Maze
; Dll ...........: kernel32.dll
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__ComUDFdll
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
; _ComListPorts
; _ComOpenPort
; _ComSetTimeouts
; _ComClosePort
;
; _ComSetBreak
; _ComClearBreak
; _ComGetInputcount
; _ComGetOutputcount
; _ComClearOutputBuffer
; _ComClearInputBuffer
;
; _ComSendByte
; _ComReadByte
; _ComSendBinary
; _ComReadBinary
;
; _ComSendChar
; _ComReadChar
; _ComSendCharArray
; _ComReadCharArray
; _ComSendString
; _ComReadString
;
; __ComClearCommError
; __PurgeComm
; ===============================================================================================================================


; ===============================================================================================================================
; Function Name:	_ComListPorts()
; Description:		Creates a list of the existing COM ports
; Parameters:       (none)
; Returns:          on success returns an array of strings
;                   on error (no com ports existing) returns -1
; Note:             The array contains e.g. "COM1", "COM2", up to "COM256"
; ===============================================================================================================================
Func _ComListPorts()
   Local $regKey = 'HKEY_LOCAL_MACHINE\HARDWARE\DEVICEMAP\SERIALCOMM'
   Local $regVal = ''
   Local $regDat = ''
   Local $comPortList[0]
   For $i = 1 To 256
	  $regVal = RegEnumVal($regKey, $i)
	  If @error Then ExitLoop
	  $regDat = RegRead($regKey, $regVal)
	  If @error Then ExitLoop
	  ReDim $comPortList[$i]
	  $comPortList[$i-1] = $regDat
   Next
   If UBound($comPortList)=0 Then Return -1
   Return $comPortList
EndFunc    ;==>_ComListPorts

; ===============================================================================================================================
; Function Name:   	_ComOpenPort($Defs)
; Description:    	Opens a serial port
; Parameters:     	$Defs    -   (string) See the notes for the format and default settings
; Returns:  		on success returns a serial port handle
;           		on error returns -1 and sets @error to 1
; Note:
; Defs is a string containing com port and many parameters. Default setting:
; COM1 baud=9600 parity=n data=8 stop=1 to=off xon=off odsr=off octs=off dtr=off rts=off idsr=off
; Usually the user would provide at least the COM port and baud settings, e.g. "COM1 baud=9600".
; All possible values are explained at https://technet.microsoft.com/en-us/library/cc732236.aspx
;
; To test the $Defs sting you can simply open a command prompt and use it with the "mode" command.
; To try e.g. "COM1 baud=9600 parity=N data=8 stop=1" just type "mode COM1 baud=9600 parity=N data=8 stop=1".
;
; The following settings aren't supported by many chipsets. If you intend to use them it is strongly
; recommended to test them with the mode command:
; data=5, data=6, parity=m, parity=s and stop=1.5
;
; References:
; Open the com port:
; CreateFile 		https://msdn.microsoft.com/en-us/library/windows/desktop/aa363858(v=vs.85).aspx
; Prepare the parameters:
; BuildCommDCB 		https://msdn.microsoft.com/en-us/library/windows/desktop/aa363143(v=vs.85).aspx
; Mode 				https://technet.microsoft.com/en-us/library/cc732236.aspx
; DCB 				https://msdn.microsoft.com/en-us/library/windows/desktop/aa363214(v=vs.85).aspx
; COMMTIMEOUTS 		https://msdn.microsoft.com/en-us/library/windows/desktop/aa363190(v=vs.85).aspx
; Set the parameters:
; SetCommState 		https://msdn.microsoft.com/en-us/library/windows/desktop/aa363436(v=vs.85).aspx
; SetCommTimeouts 	https://msdn.microsoft.com/en-us/library/windows/desktop/aa363437(v=vs.85).aspx
; ===============================================================================================================================
Func _ComOpenPort($def)
   ;The format of $def is defined here:
   ;https://msdn.microsoft.com/en-us/library/windows/desktop/aa363145(v=vs.85).aspx
   ;https://technet.microsoft.com/en-us/library/cc732236.aspx
   ;To test the $def open a command prompt windows and use it with the "mode" command
   Local Const $GENERIC_READ_WRITE = 0xC0000000
   Local Const $OPEN_EXISTING = 3
   Local Const $FILE_ATTRIBUTE_NORMAL = 0x80
   Local $hComPort, $comPort
   Local $lpFileName_struct, $lpFileName
   Local $dwDesiredAccess
   Local $dwShareMode
   Local $lpSecurityAttributes
   Local $dwCreationDisposition
   Local $dwFlagsAndAttributes
   Local $hTemplateFile

   Local $lpDef_struct, $lpDef
   Local $lpDCB_struct, $lpDCB
   Local $ReadIntervalTimeout
   Local $ReadTotalTimeoutMultiplier
   Local $ReadTotalTimeoutConstant
   Local $WriteTotalTimeoutMultiplier
   Local $WriteTotalTimeoutConstant
   Local $bool

   ;open dll
   $__ComUDFdll = DllOpen('kernel32.dll')
   If @error Then Return SetError(1, 1, -1)

   ;parse $def and set default values
   $def=StringStripWS($def, 3)	;remove leading and trailing whitespaces
												   ;default value, options
   If StringInStr($def, 'COM' )<1 Then $def =      'COM1 ' & $def  ;COM1...COM256
   If StringInStr($def, 'baud=' )<1 Then $def &=   ' baud=9600'    ;110...115200
   If StringInStr($def, 'parity=' )<1 Then $def &= ' parity=n'     ;{n|e|o|m|s} (n=none,e=even,o=odd,m=mark,s=space)
   If StringInStr($def, 'data=' )<1 Then $def &=   ' data=8'       ;{5|6|7|8}
   If StringInStr($def, 'stop=' )<1 Then $def &=   ' stop=1'       ;{1|1.5|2}
   If StringInStr($def, 'to=' )<1 Then $def &=     ' to=off'       ;{on|off}
   If StringInStr($def, 'xon=' )<1 Then $def &=    ' xon=off'      ;{on|off}
   If StringInStr($def, 'odsr=' )<1 Then $def &=   ' odsr=off'     ;{on|off}
   If StringInStr($def, 'octs=' )<1 Then $def &=   ' octs=off'     ;{on|off}
   If StringInStr($def, 'dtr=' )<1 Then $def &=    ' dtr=off'      ;{on|off|hs} (hs=handshake)
   If StringInStr($def, 'rts=' )<1 Then $def &=    ' rts=off'      ;{on|off|hs|tg} (tg=toggle)
   If StringInStr($def, 'idsr=' )<1 Then $def &=   ' idsr=off'     ;{on|off}
   $ReadIntervalTimeout=0
   $ReadTotalTimeoutMultiplier=0
   $ReadTotalTimeoutConstant=0
   $WriteTotalTimeoutMultiplier=0
   $WriteTotalTimeoutConstant=0

   ;extract $comPort from $defs
   $comPort = StringLeft($def, 4)      ;"com1".."com9" (com plus 1st digit)
   If IsNumber(StringMid($def, 5)) Then ;"com10".."com99" (2nd digit)
	  $comPort &= StringMid($def, 5)
	  If IsNumber(StringMid($def, 6)) Then ;"com100".."com256" (3rd digit)
		 $comPort &= StringMid($def, 6)
	  EndIf
   EndIf

   ;open the com port
   $lpFileName_struct = DllStructCreate('char comPort[' & StringLen($comPort)+1 & ']')
   $lpFileName = DllStructGetPtr($lpFileName_struct, 'comPort')
   DllStructSetData($lpFileName_struct, 'comPort', $comPort)

   $dwDesiredAccess = $GENERIC_READ_WRITE
   $dwShareMode = 0
   $lpSecurityAttributes = 0
   $dwCreationDisposition = $OPEN_EXISTING
   $dwFlagsAndAttributes = $FILE_ATTRIBUTE_NORMAL
   $hTemplateFile = 0
   $hComPort = DllCall($__ComUDFdll, _
	  'handle', 'CreateFile', _
	  'long_ptr', $lpFileName, _
	  'dword', $dwDesiredAccess, _
	  'dword', $dwShareMode, _
	  'long_ptr', $lpSecurityAttributes, _
	  'dword', $dwCreationDisposition, _
	  'dword', $dwFlagsAndAttributes, _
	  'handle', $hTemplateFile)
   If @error Then Return SetError(1, 1, -1)
   If UBound($hComPort) < 1 Then Return SetError(1, 1, -1)
   If Number($hComPort[0]) < 1 Then Return SetError(1, 1, -1)
   $hComPort = Number($hComPort[0])
   If @error Then Return SetError(1, 1, -1)

   ;prepare structs and pointer
   $lpDef_struct = DllStructCreate('char def[255]')
   $lpDef = DllStructGetPtr($lpDef_struct)
   DllStructSetData($lpDef_struct, 'def', $def)
   $lpDCB_struct = DllStructCreate('long DCBlength;' & _ ;DWORD DCBlength;
								 'long BaudRate;' & _  ;DWORD BaudRate
								 'long fBitFields;' & _;DWORD fBitFields
								 'short wReserved;' & _;WORD  wReserved
								 'short XonLim;' & _   ;WORD  XonLim
								 'short XoffLim;' & _  ;WORD  XoffLim
								 'byte Bytesize;' & _  ;BYTE  ByteSize
								 'byte parity;' & _    ;BYTE  Parity
								 'byte StopBits;' & _  ;BYTE  StopBits
								 'byte XonChar;' & _   ;char  XonChar
								 'byte XoffChar;' & _  ;char  XoffChar
								 'byte ErrorChar;' & _ ;char  ErrorChar
								 'byte EofChar;' & _   ;char  EofChar
								 'byte EvtChar;' & _   ;char  EvtChar
								 'short wReserved1')   ;WORD  wReserved
   $lpDCB = DllStructGetPtr($lpDCB_struct)

   ;parse $def and fill the parameters in the prepared DCB struct
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'BuildCommDCB', _
	  'long_ptr', $lpDef, _
	  'long_ptr', $lpDCB)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)

   ;set the parameters
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'SetCommState', _
	  'handle', $hComPort, _
	  'long_ptr', $lpDCB)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)

   ;set timeouts
   _ComSetTimeouts($hComPort) ; default=no timeout, return immediately

   ;return the handle to the serial port
   Return $hComPort

EndFunc   ;==>_ComOpenPort

; ===============================================================================================================================
; Function Name:	_ComSetTimeouts($hComPort,
;                                   $ReadInt = -1, $ReadMult = 0, $ReadConst = 0,
;                                   $WriteMult = 0, $WriteConst = 0)
; Description:		Could set the timeouts on the serial port
;                   It fails if non-standard values are provided, for this reason:
;                   If called with non-default settings it would break the functions
;                     to read and send data. Be warned, you don't want your program
;                     to wait forever on a missing serial port input, as the user
;                     will think the program crashed or failed.
;                     To have a stable UDF this cannot be allowed.
;                   For a background program without GUI you might want to wait
;                     for data, instead of running the idle loop/GUI loop. In this
;                     case I would recommend creating a new UDF.
;
; Parameters:       $hComPort - serial port handle (as returned by _ComOpenPort)
;                   $ReadInt - ReadIntervalTimeout
;                   $ReadMult - ReadTotalTimeoutMultiplier
;                   $ReadConst - ReadTotalTimeoutConstant
;                   $WriteMult - WriteTotalTimeoutMultiplier
;                   $WriteConst - WriteTotalTimeoutConstant
; Returns:          on success returns 0
;                   on error returns -1 and sets @error to 1
; Note:             Default setting: Return immediately from reads without waiting.
; ===============================================================================================================================
Func _ComSetTimeouts($hComPort, $ReadInt = -1, $ReadMult = 0, $ReadConst = 0, $WriteMult = 0, $WriteConst = 0)
   ;SetCommTimeouts: https://msdn.microsoft.com/en-us/library/windows/desktop/aa363437(v=vs.85).aspx
   ;COMMTIMEOUTS: https://msdn.microsoft.com/en-us/library/windows/desktop/aa363190(v=vs.85).aspx

   ;Return with an error message if non-standard parameters are provided
   ;Reasoning: Instant return is expected by all functions in this UDF
   If $ReadInt<>-1 Then Return SetError(1,1,-1)
   If $ReadMult<>0 Then Return SetError(1,1,-1)
   If $ReadConst<>0 Then Return SetError(1,1,-1)
   If $WriteMult<>0 Then Return SetError(1,1,-1)
   If $WriteConst<>0 Then Return SetError(1,1,-1)
   ;Return with an error message if non-standard parameters are provided
   ;Reasoning: Instant return is expected by all functions in this UDF

   Local $lpCommTimeouts_Struct, $lpCommTimeouts
   Local $bool
   $lpCommTimeouts_Struct = DllStructCreate( 'DWORD ReadIntervalTimeout;' & _
											 'DWORD ReadTotalTimeoutMultiplier;' & _
											 'DWORD ReadTotalTimeoutConstant;' & _
											 'DWORD WriteTotalTimeoutMultiplier;' & _
											 'DWORD WriteTotalTimeoutConstant')
   $lpCommTimeouts = DllStructGetPtr($lpCommTimeouts_Struct)
   DllStructSetData($lpCommTimeouts_Struct, 'ReadIntervalTimeout', $ReadInt)
   DllStructSetData($lpCommTimeouts_Struct, 'ReadTotalTimeoutMultiplier', $ReadMult)
   DllStructSetData($lpCommTimeouts_Struct, 'ReadTotalTimeoutConstant', $ReadConst)
   DllStructSetData($lpCommTimeouts_Struct, 'WriteTotalTimeoutMultiplier', $WriteMult)
   DllStructSetData($lpCommTimeouts_Struct, 'WriteTotalTimeoutConstant', $WriteConst)
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'SetCommTimeouts', _
	  'handle', $hComPort, _
	  'long_ptr', $lpCommTimeouts)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   Return 0
EndFunc		;==>_ComSetTimeouts

; ===============================================================================================================================
; Function Name:	_ComClosePort($hComPort)
; Description:		Closes serial port
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns 1
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComClosePort($hComPort)
   ;CloseHandle: https://msdn.microsoft.com/en-us/library/windows/desktop/ms724211(v=vs.85).aspx
   Local $bool
   $bool = DllCall($__ComUDFdll, _
	  "bool", "CloseHandle", _
	  "handle", $hComPort)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   DllClose($__ComUDFdll)
   Return 1
EndFunc   ;==>_ComClosePort



; ===============================================================================================================================
; Function Name:	_ComGetInputcount($hComPort)
; Description:		Retrieves information about the amount of available bytes in the
;                   input buffer
; Parameters:		$hComPort   - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns the amount of available bytes
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComGetInputcount($hComPort)
   Return __ComClearCommError($hComPort, 'cbInQue')
EndFunc   ;==>_ComGetInputcount

; ===============================================================================================================================
; Function Name:	_ComGetOutputcount($hComPort)
; Description:		Retrieves information about the amount of bytes in the output
;                   buffer (not yet written to the line)
; Parameters:		$hComPort   - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns the amount of bytes
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComGetOutputcount($hComPort)
   Return __ComClearCommError($hComPort, 'cbOutQue')
EndFunc   ;==>_ComGetOutputcount

; ===============================================================================================================================
; Function Name:	_ComClearOutputBuffer($hComPort)
; Description:		Discards all characters from the output buffer.
; Parameters:		$hComPort   - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns 1
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComClearOutputBuffer($hComPort)
   Return __PurgeComm($hComPort, 0x0004)
EndFunc   ;==>_ComClearOutputBuffer

; ===============================================================================================================================
; Function Name:	_ComClearInputBuffer($hComPort)
; Description:		Discards all characters from the input buffer.
; Parameters:		$hComPort   - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns 1
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComClearInputBuffer($hComPort)
   Return __PurgeComm($hComPort, 0x0008)
EndFunc   ;==>_ComClearInputBuffer

; ===============================================================================================================================
; Function Name:	_ComSetBreak($hComPort)
; Description:		Suspends character transmission for a specified communications
;                   device and places the transmission line in a break state until
;                   the _ComClearBreak function is called.
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns 1
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComSetBreak($hComPort)
   Local $bool
   ;https://msdn.microsoft.com/en-us/library/windows/desktop/aa363433(v=vs.85).aspx
   $bool = DllCall($__ComUDFdll, _
	  "bool", "ClearCommBreak", _
	  "handle", $hComPort)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   Return 1
EndFunc   ;==>_ComSetBreak

; ===============================================================================================================================
; Function Name:	_ComClearBreak($hComPort)
; Description:		Restores character transmission for a specified communications
;					device and places the transmission line in a nonbreak state.
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
; Returns:  		on success, returns 1
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComClearBreak($hComPort)
   Local $bool
   ;https://msdn.microsoft.com/en-us/library/windows/desktop/aa363179(v=vs.85).aspx
   $bool = DllCall($__ComUDFdll, _
	  "bool", "ClearCommBreak", _
	  "handle", $hComPort)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   Return 1
EndFunc   ;==>_ComClearBreak



; ===============================================================================================================================
; Function Name:	_ComSendByte($hComPort, $bSendByte)
; Description:		Send a single byte
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
;					$sSendChar - char to send
; Returns:  		on success, returns the amout of bytes written
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComSendByte($hComPort, $bSendByte)
   ;WriteFile: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365747(v=vs.85).aspx
   Local $lpBuffer_Struct, $lpBuffer
   Local $nNumberOfBytesToWrite
   Local $lpNumberOfBytesWritten_Struct, $lpNumberOfBytesWritten
   Local $lpOverlapped
   Local $bool
   $lpBuffer_Struct = DllStructCreate('byte b')
   $lpBuffer = DllStructGetPtr($lpBuffer_Struct, 'b')
   DllStructSetData($lpBuffer_Struct, 'b', $bSendByte)
   $nNumberOfBytesToWrite = 1
   $lpNumberOfBytesWritten_Struct = DllStructCreate('dword n')
   $lpNumberOfBytesWritten = DllStructGetPtr($lpNumberOfBytesWritten_Struct)
   $lpOverlapped = 0;null pointer
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'WriteFile', _
	  'handle', $hComPort, _
	  'long_ptr', $lpBuffer, _
	  'dword', $nNumberOfBytesToWrite, _
	  'long_ptr', $lpNumberOfBytesWritten, _
	  'long_ptr', $lpOverlapped)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   Return Number(DllStructGetData($lpNumberOfBytesWritten_Struct, 'n'))
EndFunc   ;==>_ComSendByte

; ===============================================================================================================================
; Function Name:    _ComReadByte($hComPort)
; Description:      Read a single byte from the com port
; Parameters:       $hComPort - serial port handle (as returned by _ComOpenPort)
; Returns:          on success returns a byte (0 to 255)
;                   on failure (empty buffer) returns -1
;                   on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComReadByte($hComPort)
   ;ReadFile: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365467(v=vs.85).aspx
   Local $lpBuffer_Struct, $lpBuffer
   Local $nNumberOfBytesToRead
   Local $nNumberOfBytesRead, $lpNumberOfBytesRead_Struct, $lpNumberOfBytesRead
   Local $lpOverlapped
   Local $byte	;received byte
   Local $bool
   $lpBuffer_Struct = DLLStructCreate('bool b')				;will contain the data read
   $lpBuffer = DllStructGetPtr($lpBuffer_Struct, 'b')
   $nNumberOfBytesToRead = 1								;number of chars to read
   $lpNumberOfBytesRead_Struct = DllStructCreate('dword n')	;will contain the number of bytes read
   $lpNumberOfBytesRead = DllStructGetPtr($lpNumberOfBytesRead_Struct)
   $lpOverlapped = 0
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'ReadFile', _
	  'handle', $hComPort, _
	  'long_ptr', $lpBuffer, _
	  'dword', $nNumberOfBytesToRead, _
	  'long_ptr', $lpNumberOfBytesRead, _
	  'long_ptr', $lpOverlapped)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   $nNumberOfBytesRead = Number(DllStructGetData($lpNumberOfBytesRead_Struct, "n"))
   if $nNumberOfBytesRead=1 then
	  $byte=DllStructGetData($lpBuffer_Struct, "b")
	  return $byte
   else
	  return -1 ;no error, no char read
   EndIf
EndFunc	;==>_ComReadByte

; ===============================================================================================================================
; Function Name:	_ComSendBinary($hComPort, $binary)
; Description:		Send binary data
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
;					$binary   - Binary data (e.g. Binary('0xFF00FF00')
; Returns:  		on success, returns the number of bytes written
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComSendBinary($hComPort, $binary)
   ;WriteFile: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365747(v=vs.85).aspx
   Local $lpBuffer_Struct, $lpBuffer
   Local $nNumberOfBytesToWrite
   Local $lpNumberOfBytesWritten_Struct, $lpNumberOfBytesWritten
   Local $lpOverlapped
   Local $bool
   $nNumberOfBytesToWrite = BinaryLen($binary)
   $lpBuffer_Struct = DllStructCreate('byte binary[' & $nNumberOfBytesToWrite & ']')
   $lpBuffer = DllStructGetPtr($lpBuffer_Struct, 'binary')
   DllStructSetData($lpBuffer_Struct, 'binary', $binary)
   $lpNumberOfBytesWritten_Struct = DllStructCreate('dword n')
   $lpNumberOfBytesWritten = DllStructGetPtr($lpNumberOfBytesWritten_Struct)
   $lpOverlapped = 0;null pointer
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'WriteFile', _
	  'handle', $hComPort, _
	  'long_ptr', $lpBuffer, _
	  'dword', $nNumberOfBytesToWrite, _
	  'long_ptr', $lpNumberOfBytesWritten, _
	  'long_ptr', $lpOverlapped)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   Return Number(DllStructGetData($lpNumberOfBytesWritten_Struct, 'n'))
EndFunc   ;==>_ComWriteChar

; ===============================================================================================================================
; Function Name:   	_ComReadBinary($hComPort, ByRef $binary, $bytesToRead)
; Description:    	Receives binary data
; Parameters:     	$hComPort - serial port handle (as returned by _ComOpenPort)
;                   $arrayBinary - will contain the read data
;				  	$bytesToRead - exact number of bytes to read
; Returns:  		on success returns 1 ($arrayBinary contains $bytesToRead bytes)
;                   on failure success 0 ($arrayBinary is undefined)
;           		on error returns -1 and sets @error to 1
; Note:
; If there are less bytes than requested in the input buffer it returns 0
; ===============================================================================================================================
Func _ComReadBinary($hComPort, ByRef $binary, $bytesToRead)
   If $bytesToRead>_ComGetInputcount($hComPort) Then Return 0
   ;ReadFile: https://msdn.microsoft.com/en-us/library/windows/desktop/aa365467(v=vs.85).aspx
   Local $lpBuffer_Struct, $lpBuffer
   Local $nNumberOfBytesToRead
   Local $nNumberOfBytesRead, $lpNumberOfBytesRead_Struct, $lpNumberOfBytesRead
   Local $lpOverlapped
   Local $byte	;received byte
   Local $bool
   $lpBuffer_Struct = DLLStructCreate('byte b[' & $bytesToRead & ']');will contain the data read
   $lpBuffer = DllStructGetPtr($lpBuffer_Struct, 'b')
   $nNumberOfBytesToRead = $bytesToRead						;number of chars to read
   $lpNumberOfBytesRead_Struct = DllStructCreate('dword n')	;will contain the number of bytes read
   $lpNumberOfBytesRead = DllStructGetPtr($lpNumberOfBytesRead_Struct)
   $lpOverlapped = 0
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'ReadFile', _
	  'handle', $hComPort, _
	  'long_ptr', $lpBuffer, _
	  'dword', $nNumberOfBytesToRead, _
	  'long_ptr', $lpNumberOfBytesRead, _
	  'long_ptr', $lpOverlapped)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   $nNumberOfBytesRead=Number(DllStructGetData($lpNumberOfBytesRead_Struct, "n"))
   If $nNumberOfBytesRead=$nNumberOfBytesToRead then
	  $binary=DllStructGetData($lpBuffer_Struct, "b")
	  Return 1
   Else
	  Return SetError(1,1,-1)
   EndIf
EndFunc   ;==>_ComReadBinary



; ===============================================================================================================================
; Function Name:	_ComSendChar($hComPort, $cChar)
; Description:		Sends an ASCII char
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
;                   $cChar    - a char
; Returns:          on success returns the number of bytes written
;                   on error returns -1 and sets @error
; Note:
; ===============================================================================================================================
Func _ComSendChar($hComPort, $cChar)
   Return _ComSendByte($hComPort, Asc($cChar))
EndFunc   ;==>_ComSendChar

; ===============================================================================================================================
; Function Name:	_ComReadChar($hComPort, $cChar)
; Description:		Reads an ASCII char
; Parameters:		$hComPort - serial port handle (as returned by _ComOpenPort)
;                   $cChar    - a char
; Returns:          on success returns a char if was read
;                   on failure (no char to read) returns -1
;                   on error returns -1 and sets @error
; Note:
; ===============================================================================================================================
Func _ComReadChar($hComPort)
   Local $bByte
   $bByte = _ComReadByte($hComPort)
   If @error Then Return SetError(1,1,-1)
   If $bByte=-1 Then
	  Return -1
   Else
	  Return Chr($bByte)
   EndIf
EndFunc   ;==>_ComReadChar

; ===============================================================================================================================
; Function Name:   	_ComSendCharArray($hComPort, $charArray)
; Description:    	Sends data
; Parameters:     	$hComPort - serial port handle (as returned by _ComOpenPort)
;					$charArray - the data to write,
; Returns:  		on success returns the number of bytes written
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComSendCharArray($hComPort, $charArray)
   Local $count
   $count = 0
   For $i = 0 To UBound($charArray)-1
	  $count += _ComSendByte($hComPort, $charArray[$i])
	  If @error Then Return SetError(1, 1, -1)
   Next
   Return $count
EndFunc   ;==>_ComSendCharArray

; ===============================================================================================================================
; Function Name:   	_ComReadCharArray($hComPort, ByRef $charArray, $charsToRead)
; Description:    	Reads data into an array
; Parameters:     	$hComPort - serial port handle (as returned by _ComOpenPort)
;                   $charArray - will contain the read data
;				  	$charsToRead - exact number of chars (bytes) to read
; Returns:  		on success returns 1 ($charArray contains $charsToRead bytes)
;           		on failure returns 0 ($charArray is undefined)
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func _ComReadCharArray($hComPort, ByRef $charArray, $charsToRead)
   Local $binary
   Local $ret
   $ret = _ComReadBinary($hComPort, $binary, $charsToRead)
   If @error Then Return SetError(1,1,-1)
   If $ret=1 Then
	  ReDim $charArray[$charsToRead]
	  For $i=0 To BinaryLen($binary)-1
		 $charArray[$i]=BinaryMid($binary,$i+1,1)
	  Next
	  Return 1
   Else
	  Return 0
   EndIf
EndFunc

; ===============================================================================================================================
; Function Name:   	_ComSendString($hComPort, $string, $encoding=1)
; Description:    	Sends data
; Parameters:     	$hComPort - serial port handle (as returned by _ComOpenPort)
;					$string   - the string to write
;					$encoding - 1=ANSI (default), 2=UTF-8
; Returns:  		on success returns the number of bytes written
;           		on error returns -1 and sets @error to 1
; Note:
; If there are less bytes than requested in the input buffer it returns 0
; ===============================================================================================================================
Func _ComSendString($hComPort, $string, $encoding=1)
   Local $charArray
   If $encoding<1 Or $encoding>2 Then Return SetError(1,1,-1)
   $charArray = StringToASCIIArray($string, 0, StringLen($string), $encoding)
   Return _ComSendCharArray($hComPort, $charArray)
EndFunc   ;==>_ComSendString

; ===============================================================================================================================
; Function Name:   	_ComReadString($hComPort, $byteLength, $encoding=1)
; Description:    	Reads a string
; Parameters:     	$hComPort   - serial port handle (as returned by _ComOpenPort)
;                   $byteLength - exact amount of bytes to read
;                                 Unicode uses more than one byte per character
;					$encoding   - 1=ANSI (default), 2=UTF-8
; Returns:  		on success returns a string
;           		on failure returns 0
;           		on error returns -1 and sets @error to 1
; Note:
; If there are less bytes than requested in the input buffer it returns 0
; ===============================================================================================================================
Func _ComReadString($hComPort, $byteLength, $encoding=1)
   Local $charArray[0]
   Local $ret
   If $encoding<1 Or $encoding>2 Then Return SetError(1,1,-1)
   $ret = _ComReadCharArray($hComPort, $charArray, $byteLength)
   If @error Then Return SetError(1,1,-1)
   If $ret=1 Then
	  Return StringFromASCIIArray($charArray, 0, UBound($charArray), $encoding)
   Else
	  Return 0
   EndIf
EndFunc

; ===============================================================================================================================
; Function Name:	__ComClearCommError($hComPort, $sInfo = 'cbInQue')
; Description:		Retrieves information about commincation errors and state
; Parameters:		$hComPort   - serial port handle (as returned by _ComOpenPort)
;					$sInfo	    - a string describing the info to return, possible values:
;	                'Errors'    - Errors that occured (see notes)
;					'Flags'	 	- Flags set or unset
;                   'cbInQue'   - The number of bytes in the input buffer (net yet read)
;                   'cbOutQue'  - The number of bytes in the output buffer (not yet send)
; Returns:  		on success, returns 0 (false), 1 (true) or a positive number
;           		on error returns -1 and sets @error to 1
; Note:
; Possible error values:
;  CE_BREAK (0x0010)    - The hardware detected a break condition
;  CE_FRAME (0x0008)    - The hardware detected a framing error
;  CE_OVERRUN (0x0002)  - A character-buffer overrun has occurred, the next
;                         character is lost
;  CE_RXOVER (0x0001)   - An input buffer overflow has occurred. There is either no
;                         room in the input buffer, or a character was received
;                         after the end-of-file (EOF) character
;  CE_RXPARITY (0x0004) - The hardware detected a parity error
; Call _ClearCommBreak to resolve errors
;
; Possible FLAG values:
; fCtsHold  (1bit) - Transmission is waiting for the CTS signal to be sent
; fDsrHold  (1bit) - Transmission is waiting for the DSR signal to be sent
; fRlsdHold (1bit) - Transmission is waiting for the RLSD signal to be sent
; fXoffHold (1bit) - Transmission is waiting because the XOFF character was received
; fXoffSent (1bit) - Transmission is waiting because the XOFF character was transmitted
; fEof      (1bit) - The end-of-file (EOF) character has been received
; fTxim     (1bit) - There is a character queued for transmission that has come to
;                    the communications device by way of the TransmitCommChar function
; Reserved (25bits)
; ===============================================================================================================================
Func __ComClearCommError($hComPort, $sInfo = 'cbInQue')
   ;https://msdn.microsoft.com/en-us/library/windows/desktop/aa363180(v=vs.85).aspx
   Local $lpErrors_struct, $lpErrors
   Local $lpStat_struct, $lpStat
   Local $status
   Local $bool
   ;prepare structs
   $lpErrors_struct = DllStructCreate('dword Errors')
   $lpErrors = DllStructGetPtr($lpErrors_struct)
   $lpStat_struct = DllStructCreate('dword Flags;' & _	;DWORD flags
									'dword cbInQue;' & _ ;DWORD cbInQue
									'dword cbOutQue;')   ;DWORD cbOutQue
   $lpStat = DllStructGetPtr($lpStat_struct)
   ;read the status
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'ClearCommError', _
	  'handle', $hComPort, _
	  'long_ptr', $lpErrors, _
	  'long_ptr', $lpStat)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   ;return requested information

   If $sInfo='cbInQue' Then ;probably most often called, checked first
	  $status=DllStructGetData($lpStat_struct, $sInfo)
	  return $status
   EndIf
   If $sInfo='cbOutQue' Then
	  $status=DllStructGetData($lpStat_struct, $sInfo)
	  return $status
   EndIf
   If $sInfo='Flags' Then
	  $status=DllStructGetData($lpStat_struct, $sInfo)
	  return $status
   EndIf
   If $sInfo='Errors' Then
	  $status=DllStructGetData($lpErrors_struct, $sInfo)
	  return $status
   EndIf
   Return SetError(1,1,-1) ;nothing matched
EndFunc   ;==>__ComClearCommError

; ===============================================================================================================================
; Function Name:	__PurgeComm($hComPort, $sFlags = 0x0000)
; Description:		Discards all characters from the output or input buffer. It can
;					also terminate pending read or write operations.
; Parameters:		$hComPort   - serial port handle (as returned by _ComOpenPort)
;					$sFlage	    - a combination of any of these numbers:
;	                0x0001      - Terminates all outstanding write operations
;					0x0002	 	- Terminates all outstanding read operations
;					0x0004		- Clears the output buffer
;                   0x0008      - Clears the input buffer
; Returns:  		on success, returns 1
;           		on error returns -1 and sets @error to 1
; Note:
; ===============================================================================================================================
Func __PurgeComm($hComPort, $sFlags = 0x0000)
   ;https://msdn.microsoft.com/en-us/library/windows/desktop/aa363428(v=vs.85).aspx
   Local $bool
   $bool = DllCall($__ComUDFdll, _
	  'bool', 'PurgeComm', _
	  'handle', $hComPort, _
	  'dword', $sFlags)
   If @error Then Return SetError(1, 1, -1)
   If $bool Then Return SetError(1, 1, -1)
   Return 1
EndFunc   ;==>__PurgeComm















