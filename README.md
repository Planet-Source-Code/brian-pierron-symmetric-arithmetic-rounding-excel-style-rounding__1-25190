<div align="center">

## Symmetric Arithmetic Rounding \(Excel style rounding\)


</div>

### Description

The Visual Basic functions CByte(), CInt(), CLng(), CCur() and Round() user a Banker's Rounding algorithm. For example, VBA.Round(0.15,1) = 0.2 **AND** VBA.Round(0.25,1) = 0.2. The following code uses Symmetric Arithmetic Rounding (similar to the Excel Worksheet Round function) where Round(0.15,1) = 0.2 and Round(0.25,1) = 0.3. Also, precision is enhanced by passing the 'Number' parameter as variant and using CDec within the routine. This helps circumvent floating point limitations. To see an excellent resource on different rounding procedures (the basis for this code) see Microsoft Article ID: Q196652.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Pierron](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-pierron.md)
**Level**          |Intermediate
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-pierron-symmetric-arithmetic-rounding-excel-style-rounding__1-25190/archive/master.zip)





### Source Code

```
Public Function Round(Number As Variant, _
           Optional NumDigitsAfterDecimal As Long) As Variant
  If Not IsNumeric(Number) Then
    Round = Number
  Else
    Round = Fix(CDec(Number * (10 ^ NumDigitsAfterDecimal)) + 0.5 * Sgn(Number)) / _
        (10 ^ NumDigitsAfterDecimal)
  End If
End Function
```

