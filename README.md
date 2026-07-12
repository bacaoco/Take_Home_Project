# Take_Home_Project
Treasury Take Home Project for AI-Powered Alcohol Label Verification App

Microsoft Access VBA SQL Queries
' These queries are used in Microsoft Access to retrieve information for multiple forms and to also generate information for multiple reports.
' The first 3 queries were used to compare record count information matched with the information displayed on the reports.
' Table - tbl_Cola in Access was a dataset retrieved from Tax and Trade Bureau, public open source website at TTB.gov/data
' Created By: Ryan Bacaoco
' Date: 7/12/2026
' Query Name: qryCola_DistilledSpirits
*************************************************************
SELECT
    *
FROM
    tbl_Cola
WHERE
    PRODUCT_TYPE = "Distilled Spirits";
*************************************************************
' Query Name: qryCola_MaltBeverages
*************************************************************
SELECT
    *
FROM
    tbl_Cola
WHERE
    PRODUCT_TYPE = "malt beverage";
*************************************************************
' Query Name: qryCola_Wines 
*************************************************************
SELECT
    *
FROM
    tbl_Cola
WHERE
    PRODUCT_TYPE = "WINE";
************************************************************
' Query Name: qry_TTB_CombinedBeverages
************************************************************
SELECT
    tbl_TTB_DistilledSpirits.BRAND_NAME,
    tbl_TTB_DistilledSpirits.ALCOHOL_CONTENT,
    tbl_TTB_DistilledSpirits.CLASS_OR_TYPE_DESIGNATION,
    tbl_TTB_DistilledSpirits.IMPORT_COUNTRY,
    tbl_TTB_DistilledSpirits.NET_CONTENTS,
    tbl_TTB_DistilledSpirits.HEALTH_WARNING_STATEMENT,
    tbl_TTB_DistilledSpirits.NAME_AND_ADDRESS,
    tbl_TTB_DistilledSpirits.COLOR_INGREDIENT_DISCLOSURE,
    tbl_TTB_DistilledSpirits.SULFITE_AND_ASPARTAME_DECLARATION,
    tbl_TTB_DistilledSpirits.NAME_AND_ADDRESS_IMPORTS,
    tbl_TTB_DistilledSpirits.SULFITE_DECLARATION,
    tbl_TTB_DistilledSpirits.APPELLATION_OF_ORIGIN,
    tbl_TTB_DistilledSpirits.PERCENTAGE_OF_FOREIGN_WINE,
    tbl_TTB_DistilledSpirits.AGE_STATEMENT,
    tbl_TTB_DistilledSpirits.COMMODITY_STATEMENT,
    tbl_TTB_DistilledSpirits.PRODUCT_TYPE
FROM
    tbl_TTB_DistilledSpirits
UNION
SELECT
    tbl_TTB_MaltBeverages.BRAND_NAME,
    tbl_TTB_MaltBeverages.ALCOHOL_CONTENT,
    tbl_TTB_MaltBeverages.CLASS_OR_TYPE_DESIGNATION,
    tbl_TTB_MaltBeverages.IMPORT_COUNTRY,
    tbl_TTB_MaltBeverages.NET_CONTENTS,
    tbl_TTB_MaltBeverages.HEALTH_WARNING_STATEMENT,
    tbl_TTB_MaltBeverages.NAME_AND_ADDRESS,
    tbl_TTB_MaltBeverages.COLOR_INGREDIENT_DISCLOSURES,
    tbl_TTB_MaltBeverages.SULFITE_AND_ASPARTAME_DECLARATION,
    tbl_TTB_MaltBeverages.NAME_AND_ADDRESS_IMPORTS,
    tbl_TTB_MaltBeverages.SULFITE_DECLARATION,
    tbl_TTB_MaltBeverages.APPELLATION_OF_ORIGIN,
    tbl_TTB_MaltBeverages.PERCENTAGE_OF_FOREIGN_WINE,
    tbl_TTB_MaltBeverages.AGE_STATEMENT,
    tbl_TTB_MaltBeverages.COMMODITY_STATEMENT,
    tbl_TTB_MaltBeverages.PRODUCT_TYPE
FROM
    tbl_TTB_MaltBeverages
UNION
SELECT
    tbl_TTB_Wines.BRAND_NAME,
    tbl_TTB_Wines.ALCOHOL_CONTENT,
    tbl_TTB_Wines.CLASS_OR_TYPE_DESIGNATION,
    tbl_TTB_Wines.IMPORT_COUNTRY,
    tbl_TTB_Wines.NET_CONTENTS,
    tbl_TTB_Wines.HEALTH_WARNING_STATEMENT,
    tbl_TTB_Wines.NAME_AND_ADDRESS,
    tbl_TTB_Wines.COLOR_INGREDIENT_DISCLOSURE,
    tbl_TTB_Wines.SULFITE_AND_ASPARTAME_DECLARATION,
    tbl_TTB_Wines.NAME_AND_ADDRESS_IMPORTS,
    tbl_TTB_Wines.SULFITE_DECLARATION,
    tbl_TTB_Wines.APPELLATION_OF_ORIGIN,
    tbl_TTB_Wines.PERCENTAGE_OF_FOREIGN_WINE,
    tbl_TTB_Wines.AGE_STATEMENT,
    tbl_TTB_Wines.COMMODITY_STATEMENT,
    tbl_TTB_Wines.PRODUCT_TYPE
FROM
    tbl_TTB_Wines;
