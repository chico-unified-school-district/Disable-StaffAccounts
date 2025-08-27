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
 EmploymentStatusCode
FROM vwHREmploymentList
WHERE
 -- Code 1,2: Certificated and Classified. Code 4: Contracted
 PersonTypeId IN (1,2,4)
 AND
 -- R: Retired T: Terminated
 EmploymentStatusCode IN ('R','T')
 AND
 EmailWork LIKE '%@%'
 AND
 -- Has to have some kind of last day listed
 ( (DateTerminationLastDay IS NOT NULL) OR (DateTermination IS NOT NULL) )
ORDER BY EmploymentStatusCode,empid