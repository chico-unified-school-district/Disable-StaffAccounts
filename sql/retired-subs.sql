SELECT
 NameFirst,
 NameLast,
 NameMiddle,
 EmailWork,
 EmailHome,
 EmpID,
 BargUnitID,
 DateTerminationLastDay,
 DateTermination,
 EmploymentStatusCode,
 EmploymentTypeDescr,
 EmploymentStatusDescr
FROM vwHREmploymentList
WHERE
 -- R: Substitute
 EmploymentStatusCode IN ('S')
 AND EmailWork LIKE '%@%'
ORDER BY EmploymentStatusCode,empid