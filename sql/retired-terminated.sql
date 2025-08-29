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
 EmploymentTypeDescr
FROM vwHREmploymentList
WHERE
 -- Code 1,2: Certificated and Classified. Code 4: Contracted
 PersonTypeId IN (1,2,4)
 -- R: Retired T: Terminated
 AND EmploymentStatusCode IN ('R','T')
 AND EmailWork LIKE '%@%'
 -- Has to have some kind of last day listed
 AND ( (DateTerminationLastDay IS NOT NULL) OR (DateTermination IS NOT NULL) )
 -- Only get recently updated rows
 AND DateTimeEdited >= DATEADD(day, -30, GETDATE())
-- ORDER BY DateTimeEdited;
ORDER BY EmploymentStatusCode,empid