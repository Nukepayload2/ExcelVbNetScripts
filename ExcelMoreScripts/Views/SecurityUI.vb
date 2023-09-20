Imports ExcelDna.Integration
Imports System.Runtime.CompilerServices
Imports System.Security.Principal
Imports System.Windows.Forms

Module SecurityUI

    Private s_canRunCode As Boolean?

    <MethodImpl(MethodImplOptions.Synchronized)>
    Function AskOrReuseCanRunCode() As Boolean
        If s_canRunCode IsNot Nothing Then
            Return s_canRunCode.Value
        End If

        Dim isElevated = New WindowsPrincipal(WindowsIdentity.GetCurrent()).
            IsInRole(WindowsBuiltInRole.Administrator)
        Dim yesButton As New TaskDialogButton With {
            .Text = My.Resources.Resources.SecurityDialog_Run,
            .ShowShieldIcon = isElevated
        }
        Dim noButton As New TaskDialogButton With {
            .Text = My.Resources.Resources.SecurityDialog_Block
        }
        Dim dlgPage As New TaskDialogPage With {
            .Caption = My.Resources.Resources.SecurityDialog_Title,
            .Buttons = New TaskDialogButtonCollection From {
                yesButton, noButton
            },
            .DefaultButton = noButton,
            .Text = My.Resources.Resources.SecurityDialog_Content
        }

        Dim buttonChoice = TaskDialog.ShowDialog(ExcelDnaUtil.WindowHandle, dlgPage)
        s_canRunCode = buttonChoice Is yesButton
        Return s_canRunCode.Value
    End Function

End Module
