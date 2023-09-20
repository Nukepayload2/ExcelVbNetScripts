Public Class DialogParentHelper

    Private Declare Unicode Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
    Private Declare Unicode Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr

    Public Shared Function ChangeParentWindow(hWndMe As IntPtr, hWndParent As IntPtr) As Boolean
        Const GWL_HWNDPARENT = -8
        Dim result As IntPtr

        If IntPtr.Size = 4 Then
            ' 32-bit system, use SetWindowLong
            result = SetWindowLong(hWndMe, GWL_HWNDPARENT, hWndParent)
        Else
            ' 64-bit system, use SetWindowLongPtr
            result = SetWindowLongPtr(hWndMe, GWL_HWNDPARENT, hWndParent)
        End If

        ' Check if the SetWindowLong/SetWindowLongPtr call was successful
        Return result <> IntPtr.Zero
    End Function

End Class
