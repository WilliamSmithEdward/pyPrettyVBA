Private Declare PtrSafe Function CopyBytes Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal dest As LongPtr, _
    ByVal src As LongPtr _
    ) As Long
