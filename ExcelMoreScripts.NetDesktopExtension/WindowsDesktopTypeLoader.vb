Imports System.Runtime.CompilerServices

Public Class WindowsDesktopTypeLoader

    Public Shared Sub LoadCommonlyUsedTypes()
        ' WinForms
        RuntimeHelpers.RunClassConstructor(GetType(System.Drawing.Bitmap).TypeHandle)
        RuntimeHelpers.RunClassConstructor(GetType(System.Windows.Forms.Form).TypeHandle)
        ' WPF
        RuntimeHelpers.RunClassConstructor(GetType(System.Windows.Window).TypeHandle)
        RuntimeHelpers.RunClassConstructor(GetType(System.Windows.UIElement).TypeHandle)
        RuntimeHelpers.RunClassConstructor(GetType(System.Windows.Rect).TypeHandle)
    End Sub

End Class
