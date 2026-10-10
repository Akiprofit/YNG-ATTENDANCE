
for birthday this month
=ARRAYFORMULA(SORT(FILTER({'Main File'!B2:B997, DAY('Main File'!C2:C997), DATEDIF('Main File'!C2:C997, TODAY(), "Y"), IF(DAY('Main File'!C2:C997) = DAY(TODAY()), "Birthday Today", IF(DAY('Main File'!C2:C997) < DAY(TODAY()), "Birthday Done", IF(DAY('Main File'!C2:C997) <= (DAY(TODAY()) + (7 - WEEKDAY(TODAY(), 2))), "Birthday This Week", "Birthday This Month")))}, 'Main File'!B2:B997 <> "", MONTH('Main File'!C2:C997) = MONTH(TODAY())), 2, TRUE))

for next month 
=ARRAYFORMULA(SORT(FILTER({'Main File'!B2:B997, DAY('Main File'!C2:C997), DATEDIF('Main File'!C2:C997, TODAY(), "Y"), IF('Main File'!B2:B997 <> "", "Birthday Next Month", "")}, 'Main File'!B2:B997 <> "", MONTH('Main File'!C2:C997) = MONTH(EDATE(TODAY(), 1))), 2, TRUE))
