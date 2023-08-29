# ExcelVbNetScripts
Provides an add-in that executes VB.NET script in Excel formulas.

## How to build
Build with .NET 6 Windows Desktop SDK or the latest Visual Studio 2022. Tested With Visual Studio 2022 version `17.6.5`.

## Example
### Run top-level code with `VB.NET.FUNCTION` function
1. Set value of A1 to the following value:
```vbnet
Dim randomId = Guid.NewGuid
Return randomId.ToString
```
2. Set formula of B1 to the following value:
```
=VB.NET.FUNCTION(A1)
```
Press Enter to run the formula function. It generates a random GUID.

### Pass parameters and run top-level code with `VB.NET.FUNCTION` function
1. Set value of A1 to the following value:
```vbnet
Return Regex.Replace(lookIn,findWhat,replacement)
```
2. Set value of A2 to the following value:
```
Excel can run Python.
```
3. Set value of A3 to the following value:
```
(?<=run )\w+
```
4. Set value of A4 to the following value:
```
VB.NET
```
5. Set formula of A5 to the following value:
```
=VB.NET.FUNCTION(A1, "lookIn", A2, "findWhat", A3, "replacement", A4)
```
Press Enter to run the formula function. It returns `Excel can run VB.NET.`.

### Run asynchronous top-level code with `VB.NET.ASYNC.FUNCTION` function
1. Set value of A1 to the following value:
```vbnet
Dim systemFolder = Environ("WinDir")
Dim winIniPath = Path.Combine(systemFolder, "win.ini")
Return Await File.ReadAllTextAsync(winIniPath)
```
2. Select B1 and enable text wrapping
3. Set formula of B1 to the following value:
```
=VB.NET.ASYNC.FUNCTION(A1)
```
Press Enter to run the formula function. It reads all text of `win.ini` asynchronously.

## Build configuration
### Default imports
```vbnet
' VB default
Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Xml.Linq
Imports System.Threading.Tasks
' Added some additional imports based on Excel built-in functions
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Numerics
```

### Options 
```vbnet
Option Compare Binary
Option Strict Off
Option Explicit On
Option Infer On
```

### Target framework
.NET 6.0 for Windows Desktop

### Referenced assemblies
All assemblies loaded by this plugin. We plan to add all assemblies in the Windows Desktop SDK.

## Special syntaxes
### Top-level code
Your code will be splitted into 2 parts, then transform into a regular VB source snippet for running. 
1. Options and imports.
2. Function body.

For example, in the following code, `Option Strict Off` and `Imports System.Net.NetworkInformation` are in the first part. Other lines are in the second part.
```vbnet
Option Strict Off
Imports System.Net.NetworkInformation

Dim pingSender As New Ping
Dim options As New PingOptions

' Set options to include the IP address and round trip time
options.DontFragment = True

' Create a buffer of 32 bytes of data to be transmitted
Dim buffer(31) As Byte

' Send the ping request
Dim reply As PingReply = Await pingSender.SendAsync(hostName, 120, buffer, options)

' Check if the ping was successful
If reply.Status = IPStatus.Success Then
    Return "Ping successful. RoundTrip time: " & reply.RoundtripTime
Else
    Return "Ping failed. Error message: " & reply.Status.ToString()
End If
```

The first part will be placed at the beginning of the template.

The second part will become the body of a function. The function will be invoked later.

### #R directive
**Important: This feature is not implemented at this time.**

#### Assembly reference
Adds assembly references in your code snippet. For example, `#R "WindowsBase.dll"` will search for the `WindowsBase.dll` in the current assembly load context and let your code snippet reference it.

#### File reference
Adds a VB source file to the project of your code snippet. For example, `#R "C:\Program.vb"` will add `C:\Program.vb` to your code snippet's project.

#### NuGet package reference
Adds a NuGet package to the project of your code snippet. For example, `#R NuGet "System.Runtime.CompilerServices.Unsafe"` will add NuGet package `System.Runtime.CompilerServices.Unsafe` to your code snippet's project.

#### Restrictions
It must be placed before any `Option` or `Imports` statements.

## Limitations
### Top-level code is based on text transformation
The current implementation of "top-level code" is not based on VB language features. It's based on text transformation in formula functions. Although it supports the `Imports` statement at the beginning of the code, it could have unexpected behavior when the code contains preprocessor directives before `Imports` statements.

Resolution:
Switch to ModVB when it supports top-level code.
