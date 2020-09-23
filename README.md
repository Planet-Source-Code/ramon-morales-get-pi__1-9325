<div align="center">

## Get  PI


</div>

### Description

Another PI program. This one allows you to get PI for any number of decimal places up to 1000. It doesn't calculate PI, it is hard coded in! All you do is pass the number of decimal places required.
 
### More Info
 
Pass a number between 1 and 1000 and the function will return PI to that many decimal places.

This is a function. Put in a code module.

Any questions contact at me at codecommandos.com or visit my web site www.codecommandos.com .

Pi with the number of decimal places that the user requested.

None (just my tired fingers ;-)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ramon\_Morales](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ramon-morales.md)
**Level**          |Beginner
**User Rating**    |3.5 (14 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[Math/ Dates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/math-dates__1-37.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ramon-morales-get-pi__1-9325/archive/master.zip)





### Source Code

```
Function GetPi(iLengthOfPI As Integer) As String
'*************************************************************
'Creation Date: 06/10/2000
'Author: Ramon Morales
'Comments: This function finds Pi to the requested number of
'decimal places. It does not calculate Pi, there is a hard
'coded string and ther result is parsed based on the requested
'number of decimal places.
'*************************************************************
Dim sPi As String
  '********************************
  'Error Trapping
  If iLengthOfPI > 1000 Or iLengthOfPI < 1 Then
    GoTo StandardExit
  End If
  '********************************
  sPi = "3.141592653589793238462643383279502884197" '
  sPi = sPi & "1693993751058209749445923078164"
  sPi = sPi & "0628620899862803482534211706798"
  sPi = sPi & "2148086513282306647093844609550"
  sPi = sPi & "5822317253594081284811174502841"
  sPi = sPi & "02701938521105559644622948954930"
  sPi = sPi & "38196442881097566593344612847564"
  sPi = sPi & "823378678316527120190914564856692"
  sPi = sPi & "346034861045432664821339360726024"
  sPi = sPi & "914127372458700660631558817488152"
  sPi = sPi & "092096282925409171536436789259036"
  sPi = sPi & "001133053054882046652138414695194"
  sPi = sPi & "151160943305727036575959195309218"
  sPi = sPi & "611738193261179310511854807446237"
  sPi = sPi & "996274956735188575272489122793818"
  sPi = sPi & "301194912983367336244065664308602"
  sPi = sPi & "139494639522473719070217986094370"
  sPi = sPi & "277053921717629317675238467481846"
  sPi = sPi & "766940513200056812714526356082778"
  sPi = sPi & "577134275778960917363717872146844"
  sPi = sPi & "0901224953430146549585371050792279"
  sPi = sPi & "6892589235420199561121290219608640"
  sPi = sPi & "3441815981362977477130996051870721"
  sPi = sPi & "1349999998372978049951059731732816"
  sPi = sPi & "0963185950244594553469083026425223"
  sPi = sPi & "0825334468503526193118817101000313"
  sPi = sPi & "7838752886587533208381420617177669"
  sPi = sPi & "1473035982534904287554687311595628"
  sPi = sPi & "6388235378759375195778185778053217"
  sPi = sPi & "1226806613001927876611195909216420"
  sPi = sPi & "1989"
StandardExit:
  On Error Resume Next
  If iLengthOfPI <= 1000 Then
    GetPi = Mid$(sPi, 1, (iLengthOfPI + 2))
  End If
  If iLengthOfPI > 1000 Then
    GetPi = "Length Too Long"
  End If
  If iLengthOfPI < 1 Then
    GetPi = "Length Must be at Least 1"
  End If
End Function
```

