# ExcelVbNetScripts
Provides an add-in that executes VB.NET script in Excel formulas.

## How to build
Build with .NET 6 Windows Desktop SDK or the latest Visual Studio 2022. Tested With Visual Studio 2022 version `17.6.5`.

## Example
### Run top-level code with `VB.NET.TOPLEVEL` function
1. Set value of A1 to the following value:
```vbnet
Dim randomId = Guid.NewGuid
Return randomId.ToString
```
2. Set formula of B1 to the following value:
```
=VB.NET.TOPLEVEL(A1)
```
Press Enter to run the formula function. It generates a random GUID.
