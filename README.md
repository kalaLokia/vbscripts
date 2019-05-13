## For MS Excel

Modules for regex based excel VBScript functions

`\b((D(-| ))?\d{4})(?=Z\b|\b)` for article without capturing Z in result

###Incentive capture
`=iferror(VLOOKUP(C:C,'C:\Users\Sabareesh\Desktop\[MONTH_DATA.xlsx]day 1'!$B:$C,2,0),iferror(VLOOKUP(C:C,'C:\Users\Sabareesh\Desktop\[MONTH_DATA2.xlsx]day 1'!$B:$C,2,0),iferror(VLOOKUP(C:C,'C:\Users\Sabareesh\Desktop\[MONTH_DATA3.xlsx]day 1'!$B:$C,2,0),iferror(VLOOKUP(C:C,'C:\Users\Sabareesh\Desktop\[MONTH_DATA4.xlsx]day 1'!$B:$C,2,0),iferror(VLOOKUP(C:C,'C:\Users\Sabareesh\Desktop\[MONTH_DATA5.xlsx]day 1'!$B:$C,2,0),iferror(VLOOKUP(C:C,'C:\Users\Sabareesh\Desktop\[MONTH_DATA6.xlsx]day 1'!$B:$C,2,0),""))))))` 
