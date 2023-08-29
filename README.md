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
Return System.Text.RegularExpressions.Regex.Replace(lookIn,findWhat,replacement)
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
