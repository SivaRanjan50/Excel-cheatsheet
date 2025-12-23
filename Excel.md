# FORMULAS & FUNCTIONS

## Essential Functions by Category
### Math & Statistical Functions

```
1. SUM
    syntax
      =SUM(range) # Adds values
    Ex:
      =SUM(A1:A10)
```

```
2. AVERAGE
    syntax
      =Average(range) # Calculates mean
    Ex:
      =AVERAGE(B2:B20)
```

```
3. COUNT
    syntax
      =COUNT(range)	 # Counts numbers
    Ex:
      =COUNT(A1:A100)
```


```
4. COUNTA
    syntax
      =COUNTA(range) # Counts non-empty cells
    Ex:
      =COUNTA(C:C)
```


```
5. MAX
    syntax
      =MAX(range) # Returns largest value
    Ex:
      =MAX(D2:D50)
```

```
6. MIN
    syntax
      =MIN(range) # Returns smallest value
    Ex:
      ==MIN(E:E)
```

```
7. MEDIAN
    syntax
      =MEDIAN(range) # Returns median value
    Ex:
      =MEDIAN(F2:F100)
```


```
8. ROUND
    syntax
      =ROUND(number, digits) # Rounds to specified digits
    Ex:
      =ROUND(3.14159, 2) → 3.14
```

```
9. SUMIF
    syntax
      =SUMIF(range, criteria, [sum_range]) # Conditional sum
    Ex:
      =SUMIF(A:A,">100",B:B)
```
```
10. SUMIFS
    syntax
      =SUMIFS(sum_range, criteria_range1, criteria1, ...)  # Multiple criteria sum
    Ex:
      =SUMIFS(C:C,A:A,">100",B:B,"<200")
```

```
11. COUNTIF
    syntax
      =COUNTIF(range, criteria) # Conditional count
    Ex:
      =COUNTIF(D:D,"Completed")
```

```
12. COUNTIFS
    syntax
      ==COUNTIFS(criteria_range1, criteria1, ...)	 # Multiple criteria count
    Ex:
      =COUNTIFS(A:A,">50",B:B,"<100")
```

### Lookup & Reference Functions

```
1. VLOOKUP
    syntax
     =VLOOKUP(lookup_value, table_array, col_index, [range_lookup]    # Vertical lookup
    Ex:
      =VLOOKUP(A2,$E$2:$G$100,3,FALSE)
```

```
2. HLOOKUP
    syntax
      =HLOOKUP(lookup_value, table_array, row_index, [range_lookup])     # Horizontal lookup
    Ex:
      =HLOOKUP("Total",A1:Z5,3,FALSE)
```

```
3. XLOOKUP
    syntax
      =XLOOKUP(lookup_value, lookup_array, return_array, [not_found], [match_mode], [search_mode])     # Modern lookup (Excel 365)
    Ex:
      =XLOOKUP(A2,B:B,C:C,"Not Found")
```

```
4. INDEX
    syntax
      =INDEX(array, row_num, [col_num])      # Returns value at position
    Ex:
      =INDEX(A1:C10,5,2)
```

```
5. MATCH
    syntax
      =MATCH(lookup_value, lookup_array, [match_type])   # Finds position of value
    Ex:
      =MATCH("Apple",A:A,0)
```

```
6. INDEX-MATCH
    syntax
      =INDEX(return_range, MATCH(lookup_value, lookup_range, 0))   # Flexible lookup combo
    Ex:
      =INDEX(C:C,MATCH(A2,B:B,0))
```

```
7. CHOOSE
    syntax
      =CHOOSE(index_num, value1, value2, ...)     # Selects from list
    Ex:
      =CHOOSE(A1,"Low","Medium","High")
```

```
8. OFFSET
    syntax
      =OFFSET(reference, rows, cols, [height], [width])     # Returns offset reference
    Ex:
      =OFFSET(A1,5,2,1,1)
```

### Text Functions

```
1. LEFT
    syntax
      =COUNTA(range)     # Counts non-empty cells
    Ex:
      =COUNTA(C:C)
```

```
2. RIGHT
    syntax
      =RIGHT(text, [num_chars])     # Extracts right characters
    Ex:
      =COUNTA(C:C)=RIGHT(A1,2)
```

```
4. MID
    syntax
      	=MID(text, start_num, num_chars)     # Extracts middle characters
    Ex:
      =MID(A1,3,2)
```

```
5. CONCAT
    syntax
      =CONCAT(text1, text2, ...)     # Joins text (Excel 2016+)
    Ex:
      =CONCAT(A1," ",B1)
```

```
6. CONCATENATE
    syntax
      =CONCATENATE(text1, text2, ...)       # Joins text (older)
    Ex:
      =CONCATENATE(A1,B1)
```

```
7. TEXTJOIN
    syntax
        =TEXTJOIN(delimiter, ignore_empty, text1, text2, ...)      # Joins with delimiter
    Ex:
      =TEXTJOIN(",",TRUE,A1:A10)
```

```
8. TRIM
    syntax
      =TRIM(text)     # Removes extra spaces
    Ex:
      =TRIM(A1)
```

```
9. UPPER
    syntax
      =UPPER(text)     # Converts to uppercase
    Ex:
      =UPPER("hello") → "HELLO"
```

```
10. LOWER
    syntax
      =LOWER(text)    #Converts to lowercase
    Ex:
      =LOWER("HELLO") → "hello"
```

```
11. PROPER
    syntax
      =PROPER(text)    #Capitalizes first letters
    Ex:
      =PROPER("john doe") → "John Doe"
```

```
12. LEN
    syntax
      =LEN(text)    # Returns text length
    Ex:
      =LEN("Excel") → 5
```

```
13. FIND
    syntax
      =FIND(find_text, within_text, [start_num])    # Finds text position
    Ex:
      =FIND("e","Excel") → 4
```

```
14. SUBSTITUTE
    syntax
      =SUBSTITUTE(text, old_text, new_text, [instance_num])     # Replaces text
    Ex:
      =SUBSTITUTE(A1,"old","new")
```

### Date & Time Functions
```
1. TODAY
    syntax
      =TODAY()     # Current date
    Ex:
      =TODAY()
```

```
2. NOW
    syntax
      =NOW()     # Current date & time
    Ex:
      =NOW()
```

```
3. DATE
    syntax
      =DATE(year, month, day)     # Creates date
    Ex:
      =DATE(2024,12,25)
```

```
4. DAY
    syntax
      =DAY(date)     # Extracts day
    Ex:
      =DAY(A1)
```

```
5. MONTH
    syntax
      =MONTH(date)     # Extracts month
    Ex:
      =MONTH(A1)
```

```
6. YEAR
    syntax
      =YEAR(date)     # Extracts year
    Ex:
      =YEAR(A1)
```

```
7. EOMONTH
    syntax
      =EOMONTH(start_date, months)     # End of month
    Ex:
      =EOMONTH(TODAY(),0)
```

```
8. WORKDAY
    syntax
      =WORKDAY(start_date, days, [holidays])     # Workday calculation
    Ex:
      =WORKDAY(A1,10)
```

```
9. NETWORKDAYS
    syntax
      =NETWORKDAYS(start_date, end_date, [holidays])     # Working days between
    Ex:
      =NETWORKDAYS(A1,B1)
```

```
10. DATEDIF
    syntax
      =DATEDIF(start_date, end_date, unit)     # Date difference
    Ex:
      =DATEDIF(A1,B1,"D")
```

```
11. WEEKDAY
    syntax
      =WEEKDAY(date, [return_type])     # Day of week
    Ex:
      =WEEKDAY(A1,2)
```

### Logical Functions

```
1. IF
    syntax
      =IF(logical_test, value_if_true, value_if_false)     # Conditional logic
    Ex:
      =IF(A1>100,"High","Low")
```

```
2. AND
    syntax
      =AND(logical1, logical2, ...)     # All conditions true	
    Ex:
      =AND(A1>0,A1<100)
```

```
3. OR
    syntax
      =OR(logical1, logical2, ...)     # Any condition true
    Ex:
      =OR(A1="Yes",B1="Yes")
```

```
4. NOT
    syntax
      =NOT(logical)     # Reverses logic
    Ex:
      =NOT(A1>100)
```

```
5. IFERROR
    syntax
      =IFERROR(value, value_if_error)     # Error handling
    Ex:
      =IFERROR(1/0,"Error")
```

```
6. IFS
    syntax
      =IFS(logical_test1, value_if_true1, ...)     # Multiple IFs (Excel 2016+)	
    Ex:
      =IFS(A1>100,"High",A1>50,"Medium",TRUE,"Low")
```

```
7. SWITCH
    syntax
      =SWITCH(expression, value1, result1, ...)     # Switch statement
    Ex:
      =SWITCH(A1,1,"One",2,"Two")
```

### Financial Functions
```
1. PMT
    syntax
      =PMT(rate, nper, pv, [fv], [type])     # Loan payment
    Ex:
      =PMT(5%/12,60,20000)
```

```
2. FV
    syntax
      =FV(rate, nper, pmt, [pv], [type])     # Future value
    Ex:
      =FV(5%/12,60,-200)
```

```
3. PV
    syntax
      =PV(rate, nper, pmt, [fv], [type])     # Present value
    Ex:
      =PV(5%/12,60,-200)
```

```
4. NPV
    syntax
      =NPV(rate, value1, value2, ...)     # Net present value
    Ex:
      =NPV(10%,B2:B10)
```

```
5. IRR
    syntax
      =IRR(values, [guess])     # Internal rate of return
    Ex:
      =IRR(B2:B10)
```

```
6. RATE
    syntax
      =RATE(nper, pmt, pv, [fv], [type], [guess])     # Interest rate
    Ex:
      =RATE(60,-200,10000)
```

### Formula Tips & Tricks
```
Absolute Reference:     $A$1   (Fixed column and row)
Mixed Reference:        $A1    (Fixed column only)
Mixed Reference:        A$1    (Fixed row only)
Relative Reference:     A1     (Changes when copied)

Named Range:           Define name in Formulas tab
                      =SUM(Sales_Data)

Array Formula:         {=SUM(A1:A10*B1:B10)} (Ctrl+Shift+Enter)

Spill Operator:        =A1# (Excel 365, references spilled range)
```

### Common Formula Errors
```
#DIV/0!    Division by zero
#N/A       Value not available
#NAME?     Unrecognized text in formula
#NULL!     Incorrect range operator
#NUM!      Problem with number
#REF!      Invalid cell reference
#VALUE!    Wrong type of argument
#######    Column too narrow
```

### Analysis Functions
```
=FORECAST.ETS()            # Exponential smoothing forecast
=TREND()                   # Linear trend line
=GROWTH()                  # Exponential growth trend
=CORREL()                  # Correlation coefficient
=SLOPE()                   # Slope of linear regression
=INTERCEPT()               # Y-intercept of regression
=LINEST()                  # Multiple linear regression
=LOGEST()                  # Exponential regression
=FREQUENCY()               # Frequency distribution
=RANK()                    # Rank of value in list
=PERCENTILE()              # Value at percentile
=QUARTILE()                # Quartile value
=STDEV()                   # Standard deviation
=VAR()                     # Variance
```


























