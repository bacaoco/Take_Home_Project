# Take_Home_Project
Treasury Take Home Project for AI-Powered Alcohol Label Verification App

Microsoft Access VBA SQL Queries
' These queries are used in Microsoft Access to retrieve information for multiple forms and to also generate information for multiple reports.
' The first 3 queries were used to compare and verify record count information matched the information displayed from the queries to the different reports.
' Table name - "tbl_Cola" is the main dataset retrieved from Tax and Trade Bureau, open source website at TTB.gov/data
' Created By: Ryan Bacaoco
' Date: 7/12/2026
*************************************************************

' Query Name: qryCola_DistilledSpirits
*************************************************************
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

*******************************************************************
*******************************************************************

' SQL and VBA Code for Access form: frmDistilledSpirits
*******************************************************************
SQL:
*******************************************************************
SELECT
    tbl_TTB_DistilledSpirits.BRAND_NAME,
    tbl_TTB_DistilledSpirits.ALCOHOL_CONTENT,
    tbl_TTB_DistilledSpirits.AGE_STATEMENT,
    tbl_TTB_DistilledSpirits.COLOR_INGREDIENT_DISCLOSURE,
    tbl_TTB_DistilledSpirits.COMMODITY_STATEMENT,
    tbl_TTB_DistilledSpirits.HEALTH_WARNING_STATEMENT,
    tbl_TTB_DistilledSpirits.NAME_AND_ADDRESS,
    tbl_TTB_DistilledSpirits.NET_CONTENTS,
    tbl_TTB_DistilledSpirits.IMPORT_COUNTRY,
    tbl_TTB_DistilledSpirits.CLASS_OR_TYPE_DESIGNATION
FROM
    tbl_TTB_DistilledSpirits;
******************************************************************
VBA CODE:
******************************************************************
Option Compare Database

Private Sub cmd_Clear_Click()
DoCmd.GoToRecord , "", acNewRec
End Sub

Private Sub cmd_Close_Click()
' Closes form and returns back to main Dasboard
DoCmd.Close
DoCmd.OpenForm ("frmDashboard")
End Sub

Private Sub cmd_DELETE_Click()
'If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbCritical, "Delete Record") = vbYes Then    
DoCmd.RunCommand acCmdSelectRecord
DoCmd.RunCommand acCmdDeleteRecord
'End If
End Sub

Private Sub cmd_Exit_Click()
DoCmd.Quit
End Sub

Private Sub cmd_Next_Click()
On Error Resume Next ' Will ignore error, if trying to go pass the last record.
DoCmd.GoToRecord , , acNext
End Sub

Private Sub cmd_Previous_Click()
On Error Resume Next ' Will ignore error, if trying to go pass the first record.
DoCmd.GoToRecord , , acPrevious
End Sub

Private Sub Form_Current()
Me.lbl_ClassType_Example.Caption = "For example,”VODKA 80-89 PROOF / VODKA”"
Me.lbl_AlcoholContectExample.Caption = "For example, ALCOHOL 40% BY VOLUME" & " or 40% Alc. by Vol."
Me.lbl_WarningStatementNotice.Caption = "***The health warning statement is the following statement, required to appear on distilled spirit labels" & _
                                        " by the Alcoholic Beverage Labeling Act (ABLA) of 1988.***"
Me.lbl_Name_Address_Example.Caption = "For example,”Bottled By: ABC Distillery Frederick, MD” is bottlers name and address statement."
Me.lbl_AutoGenerateNotice.Caption = "If Blank, click on the box to auto generate the Health Warning Statement."
Me.lbl_DistilledSpiritsLink.Caption = "https://www.ttb.gov/regulated-commodities/beverage-alcohol/distilled-spirits/labeling"
Me.lbl_DistilledSpiritsLink.HyperlinkAddress = "https://www.ttb.gov/regulated-commodities/beverage-alcohol/distilled-spirits/labeling"
End Sub

Private Sub HEALTH_WARNING_STATEMENT_Click()
' As long as the Ampere (@) sign is not used in the Format field under the Form properties. Text characters will not be restricted.
' Configure Text Format under the Data Tab and set it to Rich Text for the bold function to work for the "GOVERMENT WARNING:" statement.
HEALTH_WARNING_STATEMENT.Text = "<b>GOVERNMENT WARNING:</b> (1) According to the Surgeon General, women should not drink alcoholic beverages during pregnancy because of the risk of birth defects. 2) Consumption of alcoholic beverages impairs your ability to drive a car or operate machinery, and may cause health problems."
End Sub





