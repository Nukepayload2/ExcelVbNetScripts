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

### Run top-level code with parameters with `VB.NET.FUNCTION` function
1. Set value of A1 to the following value:
```vbnet
Return a+b
```
2. Set value of A2 to the following value:
```vbnet
100
```
3. Set value of A3 to the following value:
```vbnet
200
```
4. Set formula of B1 to the following value:
```
=VB.NET.FUNCTION(A1, "a", A2, "b", A3)
```
Press Enter to run the formula function. It returns 300.
