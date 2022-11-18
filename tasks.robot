*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF
Library             RPA.Desktop


*** Tasks ***
Insert The Sales Data For The Week And Export It As PDF
    [Documentation]    Insert The Sales Data For The Week And Export It As PDF
    Open The Intranet Website
    Download Excel Report
    Log In
    Read Excel Report
    Collect Result
    Export Table Result As PDF
    Log Out And Close Open Browser


*** Keywords ***
Open The Intranet Website
    [Documentation]    Open The Intranet Website
    Open Available Browser    https://robotsparebinindustries.com
    Maximize Browser Window

Log In
    [Documentation]    Log In
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download Excel Report
    [Documentation]    Download Excel Report
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=true

Read Excel Report
    [Documentation]    Read Excel Report
    Open Workbook    SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=true
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill And Submit Form    ${sales_rep}
    END

Fill And Submit Form
    [Documentation]    Fill And Submit Form
    [Arguments]    ${sales_rep}
    Input Text    firstname    ${sales_rep}[First Name]
    Input Text    lastname    ${sales_rep}[Last Name]
    Input Text    salesresult    ${sales_rep}[Sales]
    Select From List By Value    salestarget    ${sales_rep}[Sales Target]
    Click Button    Submit

Collect Result
    [Documentation]    Collect Result
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}sales-summary.png

Export Table Result As PDF
    [Documentation]    Export Table Result As PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_reps_result_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_reps_result_html}    ${OUTPUT_DIR}${/}sales_results.pdf

Log Out And Close Open Browser
    [Documentation]    Log Out And Close Open Browser
    Click Button    id:logout
    Close Browser
