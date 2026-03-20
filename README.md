# Excel

- Home tab font group
- Home tab alignment group
- Home tab number group
- Home tab styles group conditional
- Draw tab e cell styles
- Formulas
- Dropdown
- Pivot Tables
    - Fields
        - filters
        - columns
        - rows
        - values
    - Chart 
        - Formatting
            - Bar
            - Line
            - Area
            - Pie    
- Creating dashboard
    - data understanding
    - data transsformation
        - removing dat quality issues
            - dupicates values
            - missing values
        - new columns
            - dimension x measure
    - add charts
    - slicers
        - formatting
    - report connections
- Power Query Editor
    - Data
        - from table
            - column distribution
                - unique value
                - null value
            - column profile
            - column quality
    - Close and Load
- Combine
    - Merge
        - inner join
            - data
            - from trable
            - close and load to...
            - only create connection
            - data
            - from table
            - merge queries
            - merge ueries as new
            - join kind
                - inner
            - expand table column
        - left outer join
            - get data
            - launch power query editor
            - merge querie as new
            - join kind
                - left outer

        - right outer join
            - get data
            - launch power query editor
            - merge querie as new
            - join kind
                - right outer
        - left anti join
        - right anti join
        - full outer join
    - append
        - twoor more tables
    - Home
        - split column
        - group
        - choose column
        - remove column
        - format
    - Transform
    - add column
        - conditional coumn

- Referencing
    - relative reference: A1
    - absolute reference $A$1
    - mixed reference: $A1, A$1

- shortcut
    - f2: edit cell
    - f4: freeze cell or range
    - ctrl+1: format window
    - ctrl+down arrow: last cell
    - ctrl+up arrow: first cell
    - ctrl+semi colon: current date
    - ctrl+shift+semi colon: current time
    - ctrl+space: select entire colum
    - shift+space:select entire row
    - ctrl+minus: delete row or colum per selection
    - ctrl+shift+plus: add row or coumn
    - ctrl+shift+down arrow: down selection
    - ctrl+shift+up arrow: up selection
    - ctrl+shift+right arrow: right selection
    - ctrl+shift+left arrow: left selection
    - ctrl+shift+l: apply filter
    - shift+f11: add newsheet

- format painter

- paste special

# Data Extraction

- =Vlookup(Lookup_Value, Table_Array, Column Index number , 0)
    - example: 
        - VLOOKUP(MIN(Dataset2!$K$2:$K$8400), Dataset2!$K$1:$U$84000, 8, 0)
        - VLOOKUP(MAX(Dataset2!$K$2:$K$8400), Dataset2!$K$1:$U$84000, 8, 0)
        - VLOOKUP(LARGE(Dataset2!$K$2:$K$8400, ROW(Sheet2!J1)), Dataset2!$K$1:$U$84000, 10, 0)
        - VLOOKUP(SMALL(Dataset2!$K$2:$K$8400, ROW(Sheet2!J1)), Dataset2!$K$1:$U$84000, 10, 0)

- =Match(lookup_value, lookup_array, 0)
    - =MATCH($A2, Dataset1!$B1:$B50, 0)

- =vlookup($a2, choose({1,2}, Dataset1!$B1:B50, Dataset1!$A1:$A50), 2, 0)

- volookup($l2, Dataset1!$B$1:$U$400, MATCH(M$1, Dataset1!$B$1:$U$1,0), 0)

- VOLOOKUP($F5, $F$12:$G$18, 2, 0)

- VOLOOKUP(VOLOOKUP($F5, $F$12:$G$18, 2, 0), $F$20:$G$26, 2, 0)

- 
```
=IFERROR(
   IFERROR(
      IFERROR(
         VLOOKUP($C12,$A$1:$B$7,MATCH(G$11,$A$1:$B$1,0),0),
         VLOOKUP($C12,$D$1:$E$7,MATCH(G$11,$D$1:$E$1,0),0)
      ),
      VLOOKUP($C12,$G$1:$H$7,MATCH(G$11,$G$1:$H$1,0),0)
   ),
   VLOOKUP($C12,$J$1:$K$7,MATCH(G$11,$J$1:$K$1,0),0)
)
```

- 
```
=IFERROR(
    IFERROR(
        IFERROR(
            VLOOKUP(
                $C25, $A$2:$H$19, MATCH(D$24, $A$2:$H$2, 0), 0
            ),
            VLOOKUP(
                $C25, $J$2:$Q$19, MATCH(D$24, $J$2:$Q$2, 0), 0                
            )
        ),
        VLOOKUP(
            $C25, $S$2:$Z$17, MATCH(D$24, $S$2:$Z$2, 0), 0              
        )
    ),
    "No Data"
)
```

- 
```
=VLOOKUP(
    $H25&"*", $AT$1:$AU$6, 2, 0
)
```

-
```
=HLOOKUP(
    LOOKUP_VALUE, TABLE_ARRAY, ROW INDEX NUMBER, 0
)
```

- 
```
=HLOOKUP(
    C$14, Dataset1!$A$1>$H$50, 9, 0
)
```

- 
```
=HLOOKUP(
    F$14, Data1!$A$1:$H$50, MATCH('Func DataExraction'!$E15, DATA1!$B$1:$B$50, 0), 0
)
```

- 
```
=INDEX(Data1!$A$1:$H$50,
       MATCH('Func DataExraction'!$E15, Data1!$B$1:$B$50,0),
       MATCH(F$14, Data1!$A$1:$H$1,0)
)
```

- 
```
= iferror(
    iferror(
        hlookup(
            b$1, sheet1!$a$1:$b$8, match(sheet2!$a2, sheet1!$a$1:$a$8, 0), 0
        ),
        hlookup(
            b$1, sheet!$d$1:$e$8, match(sheet2!$a2, sheet!$d$1:$d$8, 0), 0
        )
    ),
    hlookup(
        b$1, sheet!$g$1:$h$8, match(sheet2!$a2, sheet!$g$1:$g$8, 0), 0
    )
)
```

- 
```
=IFERROR(
   IFERROR(
      INDEX(Sheet1!$A$1:$B$8,
            MATCH(Sheet2!$A2, Sheet1!$A$1:$A$8,0),
            MATCH(B$1, Sheet1!$A$1:$B$1,0)
      ),
      INDEX(Sheet1!$D$1:$E$8,
            MATCH(Sheet2!$A2, Sheet1!$D$1:$D$8,0),
            MATCH(B$1, Sheet1!$D$1:$E$1,0)
      )
   ),
   INDEX(Sheet1!$G$1:$H$8,
         MATCH(Sheet2!$A2, Sheet1!$G$1:$G$8,0),
         MATCH(B$1, Sheet1!$G$1:$H$1,0)
   )
)
```

- 
```
=Index(DatabaseRange, Row, Column)
=Index(DatabaeRange, Match(lookupovalue, databaserange, 0), match(lookup_value, databarange, 0))
```



- 
```
=INDEX(Dataset1!$a$1:$h$50, 35, 3)
```

- 
```
=INDEX(
    Dataset1!$a$1:$h$50, match(sheet1!$a$8:$a$13, dataset1!$b$1:$b50, 0), 1
)
```

-
```
=INDEX(
    Dataset1!$A$1:$H$50,
    MATCH(Sheet1!$E8, Dataset1!$B$1:$B$50, 0),
    MATCH(F$7, Dataset1!$A$1:$H$1, 0)
)
```

- 
```
=IFERROR(
    INDEX(Dataset1!$A$1:$U$400,
          MATCH(Sheet1!$A2, Dataset1!$C$1:$C$400, 0),
          MATCH(Sheet1!B$1, Dataset1!$A$1:$U$1, 0)
    ),
    INDEX(Dataset2!$A$1:$U$401,
          MATCH(Sheet1!$A2, Dataset2!$C$1:$C$401, 0),
          MATCH(Sheet1!B$1, Dataset2!$A$1:$U$1, 0)
    )
)
```

# Data Cleaning

```
Len(Cell reference)
```

```
Find(Finda text, within text, start number)
```

```
=find("K", "RAM KUMAR", 2)
```
```
=left(A2, FIND(" ", A2))
```
```
RIGHT(A2, LEN()A2-FIND(" ", A2))
```
```
=FIND("a", "Amar", 1), output: 3 because it is case sensitive
```
```
=SEARCH("K", A9)
```

```
=LEFT(A1, 4), OUT PUT: RAJE
```
```
=LEF(A1, FIND(" ", A1)), OUT PUT: RAJESH
```
```
=LEFT(A1, SEARCH(" ", A1)) OUTPUT RAJESH
```

```
=FIND(" ", A1)
```

```
=LEFT(A1, FIND(" ", A1))
```

```
=TRIM(LEFT(A1, FIND(" ", A1)))
```

```
=RIGHT(A1, 4)
```

```
=RIGHT(A1, LEN(A1)-FIND(" ", A1))
```

```
=RIGHT(A1, SEARCH(" ", A1))
```

```
=MID(TEXT, START NUMBER, NUMBER CHARACTER)
```

```
=MID("TRAINING", 4, 2) OUTPUT: "IN"
```

```
=MID(A2, SEARCH(" ", A2, 1)+1, SEARCH(" ", A2, SEARCH(" ", A2, 1)+1)-SEARCH(" ", A2, 1)) OUTPUT:KUMAR
```

```
=MID(A1, FIND(" ", A1)+1, FIND(" ", A1, FIND(" ", A1)+1)-FIND(" ", A1)-1)
```

```
SUBSTITUTE(A1, " ", "*", 2)
```

```
=substitute(
    substitute(
        substitute(
            B1, "(", ""
        ),
        "-",""
    ),
    "!",""
)
```

```
=len(a1)-len(substitute(a1, "o"))
```

```
=substitute(a1, " ", "*", len(a1)-len(substitute(a1, " ", "")))
```
```
=REPLACE(old text, start, number of chaacters, new text))
```

```
=REPLACE(A32, FIND(" ", A32)+1, 5, "jha")
```

```
=LEFT(REPLACE(A37,1,FIND("Invoice Number",A37)+14,""),FIND(" ",REPLACE(A37,1,FIND("Invoice Number",A37)+14,"")))
```

```
=RIGHT(A19,LEN(A19)-FIND("*",SUBSTITUTE(A19," ","*",LEN(A19)-LEN(SUBSTITUTE(A19," ","")))))
```

# Dynamic range with offset function

```
=OFFSET(A3, 1, 1)
```

```
=average(offset(a3,1,1,45,3))
```


```
=SUM(OFFSET(A3, 1, 1, COUNTA(A:A)-1, COUNTA(10:10)))
```

- Create a dynamic chart with offest function
```
select dataset
insert
chart layout
bar
```

```
=VLOOKUP($A10,OFFSET(Data1!$B$1,0,0,COUNTA(Data1!B:B),7),2,0)
```

- Create dynamic pivot table with offeset

# Data aggregation

```
=SUMPRODUCT($B$2:$B$8400, $C$2:$C$8400)
```
```
=COUNTIF($A$2:$A$50, A2)
```
```
=IF(COUNTIF($A$2>$A$50, A2)>1, 'DUP','UNIQUE')
```
```
=SUMPRODUCT(--(Dataset1!$C$2:$C$50=Report!A2))
```
```
=COUNTIF($K$4:$K$13,'Mrs.")
```

```
=COUNTIFS(Dataset1!$E$2:$E$50,Summary!$A2,Dataset1!$C$2:$C$50,Summary!B$1)
```

```
=SUMPRODUCT((Dataset1!$E$2:$E$50=Summary!$F2)*(Dataset1!$C$2:$C$50=Summary!G$1))
```

```
=COUNTIFS(Dataset1!$C$2:$C$50,Summary!$A13,Dataset1!$D$2:$D$50,">20",Dataset1!$D$2:$D$50,"<=40")
```

```
=SUMPRODUCT((Dataset1!$C$2:$C$50=Summary!$F13)*(Dataset1!$D$2:$D$50>20)*(Dataset1!$D$2:$D$50<=40))
```
```
=COUNTIFS(Dataset3!$A$2:$A$36865,Summary!$A18,Dataset3!$K$2:$K$36865,Summary!B$17)
```
```
=SUMPRODUCT((Dataset3!$A$2:$A$36865=Summary!$F18)*(Dataset3!$K$2:$K$36865=Summary!G$17))
```
```
=COUNTIFS(Dataset2!$N$2:$N$8400,Summary!$A28,Dataset2!$P$2:$P$8400,Summary!B$27)
```

```
=SUMPRODUCT((Dataset2!$N$2:$N$8400=Summary!$F28)*(Dataset2!$P$2:$P$8400=Summary!G$27))
```
```
=SUMPRODUCT((Dataset2!$N$2:$N$8400=Summary!$F28)*(Dataset2!$P$2:$P$8400=Summary!G$27)*(TEXT(Dataset2!$C$2:$C$8400,"YYYY")=Summary!$F$26))
```

```
=SUMIF(Dataset2!$N$2:$N$8400,Summary!$A5,Dataset2!$F$2:$F$8400)
```

```
=SUMPRODUCT((Dataset2!$N$2:$N$8400=Summary!$A5)*(Dataset2!$F$2:$F$8400))
```

```
=SUMIF(Dataset2!$N$2:$N$8400,Summary!$A19,Dataset2!$I$2:$I$8400)
```

```
=SUMPRODUCT((Dataset2!$N$2:$N$8400=Summary!$A19)*(Dataset2!$I$2:$I$8400))
```

```
=SUMIFS(Dataset2!$F$2:$F$8400,Dataset2!$N$2:$N$8400,Summary!$A6,Dataset2!$P$2:$P$8400,Summary!B$5)
```

```
=SUMIFS(Dataset2!$I$2:$I$8400,Dataset2!$N$2:$N$8400,Summary!$A26,Dataset2!$P$2:$P$8400,Summary!D$19)
```

```
=IF($A$18="Profit",SUMIFS(Dataset2!$I$2:$I$8400,Dataset2!$N$2:$N$8400,Summary!$A20,Dataset2!$P$2:$P$8400,Summary!B$19),SUMIFS(Dataset2!$J$2:$J$8400,Dataset2!$N$2:$N$8400,Summary!$A20,Dataset2!$P$2:$P$8400,Summary!B$19))
```

```
=SUMPRODUCT((Dataset2!$F$2:$F$8400)*(Dataset2!$N$2:$N$8400=Summary!$F10)*(Dataset2!$P$2:$P$8400=Summary!G$5))
```

```
=IF($A$18="Profit",
SUMIFS(Dataset2!$I$2:$I$8400,Dataset2!$N$2:$N$8400,Summary!$A20,Dataset2!$P$2:$P$8400,Summary!C$19),
SUMIFS(Dataset2!$J$2:$J$8400,Dataset2!$N$2:$N$8400,Summary!$A20,Dataset2!$P$2:$P$8400,Summary!C$19,Dataset2!$I$2:$I$8400,"<0"))
```

```
=AVERAGEIF(Dataset2!$N$2:$N$8400,Practice1!$A5,Dataset2!$F$2:$F$8400)
```

```
=AVERAGEIF(Dataset2!$N$2:$N$8400,Practice1!$A5,Dataset2!$I$2:$I$8400)
```

```
=AVERAGEIFS(Dataset2!$F$2:$F$8400,Dataset2!$N$2:$N$8400,Practice1!$F5,Dataset2!$P$2:$P$8400,Practice1!G$4)
```

```
=IF($C$3="Countifs",
COUNTIFS(Dataset1!$E$2:$E$50,Summary!$C5,Dataset1!$C$2:$C$50,Summary!D$4),
IF($C$3="Averageifs",
AVERAGEIFS(Dataset1!$J$2:$J$50,Dataset1!$E$2:$E$50,Summary!$C5,Dataset1!$C$2:$C$50,Summary!D$4),
IF($C$3="Sumifs",
SUMIFS(Dataset1!$J$2:$J$50,Dataset1!$C$2:$C$50,Summary!D$4,Dataset1!$E$2:$E$50,Summary!$C5))))
```

```
=SUMPRODUCT(($C$10:$C$61=$AU$12)*($E$10:$AP$61))
```

```
=SUMPRODUCT(($C$10:$C$61=$AU$12)*($E$10:$AP$61))
```

```
=SUMPRODUCT(($C$10:$C$61=$AU$12)*($E$10:$AP$61=AV$11))
```

# Time Series Analysis

```
=day("05/04/2014")
```

```
=month("05/04/2014")
```

```
=year("05/04/2014")
```

```
=date(year, moth, day)
```


```
=today()
```

```
=now()
```

```
=hour(now())
```

```
=minute(now())
```

```
=second(now())
```

```
=edate(now(), 3)
```

```
=eomonth()
```

```
=text(a8, "DD")
```

```
=text(a8, "DDD")
```

```
=text(a8, "DDDD")
```

```
=text(a8, "MM")
```

```
=text(a8, "MMM")
```

```
=text(a8, "MMMM")
```

```
=NETWORKDAYS(A1, A30, A8:A9)
```

```
=NETWORKDAYS.INTL(A1, A38)
```

```
WORKDAY
```

```
WORKDAY.INTL
```

```
WEEKDAY
```

```
WEEKNUM
```

```
DATEDIFF
```


# VBA Training

Standards for visual baisc application. VBA é programming language to use automate manual task in any office applications. We have to write script which is called macro. Macro is created in module, sheet, class module or workbook editor, which is called VBE.

- Developer
    - Visual basic
        - insert
            - module
```
Sub Basic()
    Sheets.Add after:=Sheets(Sheets.Count)
    Range("A1") = "Name"
    Range("B1") = "Salary"

    Range("A1:B1").Font.Bold = True

    Range("A2") = "John"
    Range("A3") = "Akash"

    Range("B2") = 56000
    Range("B3") = 34000

    Range("C1") = "Tax"
    Range("C2") = "B2*.10"

    Range("C2").AutoFill Range("C2:C3")

    Range("B2").Interior.Color = vbGreen

End Sub
```




# Analyst

Analyst analyze infrmtion, data, or processes to provide insights, recommendations, or solutions. Analystis are often specialists in a particular area and use their skills to inerpret and make sense of complex information.

# Data Analyst

Examines and interprets data to provide meaningful insights. Data analysts ue statical techniques and data visualization tools to analyze and present data in a way that helps organizations make informed decisions.


        
