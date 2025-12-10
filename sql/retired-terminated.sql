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
 -- R: Retired T: Terminated
 EmploymentStatusCode IN ('R','T')
 -- Code 1,2: Certificated and Classified. Code 4: Contracted
 AND PersonTypeId IN (1,2,4)
 AND EmailWork LIKE '%@%'
 -- Has to have some kind of last day listed
 -- AND ( (DateTerminationLastDay IS NOT NULL) OR (DateTermination IS NOT NULL) )
 -- Only get recently updated rows
 -- AND DateTimeEdited >= DATEADD(day, -30, GETDATE())
-- ORDER BY DateTimeEdited;
ORDER BY EmploymentStatusCode,empid