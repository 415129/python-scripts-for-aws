WITH
  base_data AS (
    SELECT
      emp.EMPLOYEE,
      lce.DEFAULT_FLAG,
      lce.ACH_DIST_NBR
    FROM
      `prj-pvt-oneerp-data-raw-78c9.lawson_chi.employee` emp
    LEFT JOIN
      `prj-pvt-oneerp-data-raw-78c9.lawson_chi.emachdepst` lce
      ON emp.EMPLOYEE = lce.EMPLOYEE AND lce.END_DATE IS NULL
    WHERE
      emp.EMP_STATUS NOT IN ('C1', 'C2', 'T2', 'W2', 'S1')
  ),
  ranked_data AS (
    SELECT
      EMPLOYEE,
      DEFAULT_FLAG,
      ACH_DIST_NBR,
      COUNT(ACH_DIST_NBR) OVER (PARTITION BY EMPLOYEE) AS distribution_count,
      MAX(ACH_DIST_NBR) OVER (PARTITION BY EMPLOYEE) AS SUPPLEMENTAL_PAYMENTS
    FROM
      base_data
  )
SELECT
  r.EMPLOYEE AS Worker_Reference_ID,
  'USA' AS Worker_Country_Reference_ID,
  'USD' AS Worker_Currency_Reference_ID,
  r.DEFAULT_FLAG,
  'REGULAR_PAYMENTS' AS Payment_Election_Rule_ID,
  CASE
    WHEN r.DEFAULT_FLAG = 'Y' THEN r.distribution_count
    ELSE r.ACH_DIST_NBR
  END AS Payment_Election_Order,
  'Direct_Deposit' AS Payment_Type_Reference_ID,
  'USA' AS Bank_Account_Country_Reference_ID,
  'USD' AS Bank_Account_Currency_Reference_ID
FROM
  ranked_data r
UNION ALL
SELECT
  r.EMPLOYEE AS Worker_Reference_ID,
  'USA' AS Worker_Country_Reference_ID,
  'USD' AS Worker_Currency_Reference_ID,
  r.DEFAULT_FLAG,
  'SUPPLEMENTAL_PAYMENTS' AS Payment_Election_Rule_ID,
  r.SUPPLEMENTAL_PAYMENTS AS Payment_Election_Order,
  'Direct_Deposit' AS Payment_Type_Reference_ID,
  'USA' AS Bank_Account_Country_Reference_ID,
  'USD' AS Bank_Account_Currency_Reference_ID
FROM
  ranked_data r
WHERE
  r.DEFAULT_FLAG = 'Y';