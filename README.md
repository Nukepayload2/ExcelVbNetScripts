# ExcelVbNetScripts
Provides an add-in that executes VB.NET script in Excel formulas.

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
