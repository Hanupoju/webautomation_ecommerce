*** Settings ***
Library    SeleniumLibrary    
Library    SikuliLibrary        
Library    Excel.Application
Library    String    
Library    Collections   
# Suite Setup    Sikuli Startup

*** Variables ***
${AmazonURL}    https://www.amazon.in/
${FlipkartURL}    https://www.flipkart.com/
${AmazonSearchbox}    //input[@id="twotabsearchtextbox"]
${FlipkartSearchbox}    //input[@title='Search for products, brands and more']
${iOSCheckbox}    //span[text()='iOS']
${AppleCheckbox}    //span[text()='Apple']
${FeatureDropdown}    //span[text()='Featured']
${AmazonPrAttribute}    //div[@class='s-include-content-margin s-border-bottom s-latency-cf-section']
${FlipkartPrAttribute}    //div[@class='MIXNux']
${FlipkartPrPriceAttribute}    //div[@class='_30jeq3 _1_WHN1']
${NoOfPages}    //a[text()='Next']//parent::li//preceding-sibling::li[1]
${NoOfPagesFlip}    //span[text()='Next']/parent::a/preceding-sibling::a[1]
${NextPageBtn}    //a[text()='Next']
${AmazonPrNameAttribute}    //span[@class='a-size-medium a-color-base a-text-normal']
${FlipkartPrNameAttribute}    //div[@class='_4rR01T']
${AmazonPrLinkAttribute}    //a[@class='a-link-normal a-text-normal']
${FlipkartPrLinkAttribute}    //a[@class='_1fQZEK']
${FlipkartLoginpopup}    //a[text()='New to Flipkart? Create an account']
${NextPageBtnFlip}    //span[text()='Next']
${Obj_Minimum_Timeout}    15
${items_in_single_page}    None
${ObjectCount}    None
${addedVal}    0
${rcnt}    0
${items_in_page}    0

*** Test cases ***
Scenario
    
	Launching web application1
    Launching web application2
    Excel Price Sorting

*** Keywords ***

Launching web application1
    
    Open Browser    ${AmazonURL}    CHROME
    Maximize Browser Window
    Wait Until Page Contains Element    ${AmazonSearchbox}    ${Obj_Minimum_Timeout}
    SeleniumLibrary.Input Text    ${AmazonSearchbox}    iPhone 11
    Press Keys    none    RETURN
    Sleep    3
    
    Wait Until Page Contains Element    ${iOSCheckbox}    ${Obj_Minimum_Timeout}
    Click Element    ${iOSCheckbox}
    ${Stts}    Run Keyword And Return Status    Page Should Contain Element    ${AppleCheckbox}    ${Obj_Minimum_Timeout}  
    Run Keyword If    ${Stts} == ${True}    Click Element    ${AppleCheckbox}     
    Sleep    4
    Wait Until Page Contains Element    ${AmazonPrAttribute}    ${Obj_Minimum_Timeout}    
    Wait Until Page Contains Element    ${NoOfPages}    ${Obj_Minimum_Timeout}
    ${TotalPages}    SeleniumLibrary.Get Text    ${NoOfPages}
    ${items_in_page}    Get Element Count    ${AmazonPrAttribute}
    Set Suite Variable    ${items_in_page}
    Excel Headers
    
    FOR    ${j}    IN RANGE    1    ${TotalPages}+1
        ${items_in_single_page}    Get Element Count    ${AmazonPrAttribute}
        Set Suite Variable    ${items_in_single_page} 
        ${n}    evaluate    ${j}-1
        ${addedVal}    evaluate    ${items_in_page}*${n}
        Set Suite Variable    ${addedVal}
        Get Product Properties
        ${status}    Run Keyword And Return Status    Page Should Contain Element    ${NextPageBtn}    ${Obj_Minimum_Timeout}  
        Run Keyword If    ${status} == ${True}    Click Element    ${NextPageBtn}
        Sleep    3
    END
    Quit Application    

Get Product Properties
    
    FOR    ${i}    IN RANGE    1    ${items_in_single_page}+1
        Wait Until Page Contains Element    (${AmazonPrNameAttribute})[${i}]    ${Obj_Minimum_Timeout}
        ${PrName}    SeleniumLibrary.Get Text    (${AmazonPrNameAttribute})[${i}]
        Wait Until Page Contains Element    (${AmazonPrLinkAttribute})[${i}]    ${Obj_Minimum_Timeout}
        ${PrLink}    Get Element Attribute    (${AmazonPrLinkAttribute})[${i}]    href
        ${link}    Get Substring    ${PrLink}    22    None

        Wait Until Page Contains Element    (//a[@href='/${link}'])[last()]/child::span    ${Obj_Minimum_Timeout}
        ${Price}    SeleniumLibrary.Get Text    (//a[@href='/${link}'])[last()]/child::span
        ${PrPrice}    Get Substring    ${Price}    1    None   

        Excel.Application.Set Active Worksheet    Sheet1
        ${rcnt}    Evaluate    ${addedVal}+${i}    
        ${row}    Evaluate    ${rcnt}+1    
        Write To Cells    row= ${row}    column= 1    value=${PrName}
        Write To Cells    row= ${row}    column= 2    value=${PrPrice}
        Write To Cells    row= ${row}    column= 3    value=Amazon
        Write To Cells    row= ${row}    column= 4    value=${Prlink}+${SPACE}
        Set Suite Variable    ${rcnt}
        
    END
    Save Excel
    
Launching web application2
    
    Open Browser    ${FlipkartURL}    CHROME
    # Wait Until Screen Contain    Uname    ${Obj_Minimum_Timeout}
    # SikuliLibrary.Input Text    Uname    z018350
    # SikuliLibrary.Input Text    Pass    change99
    # Press Special Key    ENTER
    Maximize Browser Window
    
    Wait Until Page Contains Element    ${FlipkartLoginpopup}    ${Obj_Minimum_Timeout}
    Press Keys    none    ESCAPE
        
    Wait Until Page Contains Element    ${FlipkartSearchbox}    ${Obj_Minimum_Timeout}

    SeleniumLibrary.Input Text    ${FlipkartSearchbox}    iPhone 11
    Press Keys    none    RETURN
    
    Sleep    3
    Wait Until Page Contains Element    ${NextPageBtnFlip}    ${Obj_Minimum_Timeout}
    ${TotalPages2}    SeleniumLibrary.Get Text    ${NoOfPagesFlip}
    ${ObjectCnt}    Get Element Count   ${FlipkartPrAttribute}
    Excel.Application.Open Application    1    1
    Excel.Application.Open Workbook    ${CURDIR}\\WebAutomationTask\\..\\Output.xlsx
    FOR    ${l}    IN RANGE    1    ${TotalPages2}+1
        Wait Until Page Contains Element    ${FlipkartPrAttribute}    ${Obj_Minimum_Timeout}
        ${ObjectCount}    Get Element Count   ${FlipkartPrAttribute}
        Set Suite Variable    ${ObjectCount} 
        ${n}    evaluate    ${l}-1
        ${addedVal}    evaluate    ${ObjectCnt}*${n}
        Set Suite Variable    ${addedVal} 
        Get Product Properties 2
        ${status}    Run Keyword And Return Status    Page Should Contain Element    ${NextPageBtnFlip}    ${Obj_Minimum_Timeout}  
        Run Keyword If    ${status} == ${True}    Click Element    ${NextPageBtnFlip}       
        Sleep    3
    END
    Excel.Application.Quit Application    
    Close All Browsers
    
Get Product Properties 2
    
    FOR    ${k}    IN RANGE    1    ${ObjectCount}+1
        Wait Until Page Contains Element    (${FlipkartPrNameAttribute})[${k}]    ${Obj_Minimum_Timeout}
        ${PrName2}    SeleniumLibrary.Get Text    (${FlipkartPrNameAttribute})[${k}]
        Wait Until Page Contains Element    (${FlipkartPrLinkAttribute})[${k}]    ${Obj_Minimum_Timeout}
        ${PrLink2}    Get Element Attribute    (${FlipkartPrLinkAttribute})[${k}]    href
        ${link2}    Get Substring    ${PrLink2}    19    None

        Wait Until Page Contains Element    (${FlipkartPrPriceAttribute})[${k}]
        ${Price2}    SeleniumLibrary.Get Text    (${FlipkartPrPriceAttribute})[${k}]
        ${PrPrice2}    Get Substring    ${Price2}    1    None   
        
        Excel.Application.Set Active Worksheet    Sheet1
        ${rcnt2}    Evaluate    ${rcnt}+${addedVal}+${k}
        ${row}    Evaluate    ${rcnt2}+1    
        Write To Cells    row= ${row}    column= 1    value=${PrName2}
        Write To Cells    row= ${row}    column= 2    value=${PrPrice2}
        Write To Cells    row= ${row}    column= 3    value=Flipkart
        Write To Cells    row= ${row}    column= 4    value=${PrLink2}+${SPACE}
    END
    Save Excel
        
Excel Headers
    
    Excel.Application.Open Application    1    1
    Excel.Application.Open Workbook    ${CURDIR}\\WebAutomationTask\\..\\Output.xlsx
    Excel.Application.Set Active Worksheet    Sheet1
    
    Write To Cells    row= 1    column= 1    value=Product Name
    Write To Cells    row= 1    column= 2    value=Product price
    Write To Cells    row= 1    column= 3    value=Online Store
    Write To Cells    row= 1    column= 4    value=Product Link
   
Excel Price Sorting
    
    Sleep    2
    Excel.Application.Open Application    1    1
    Excel.Application.Open Workbook    ${CURDIR}\\WebAutomationTask\\..\\Output.xlsx
    Excel.Application.Set Active Worksheet    Sheet1
    Save Excel
    Sleep    7
    Repeat Keyword    10    Press Special Key    LEFT
    Repeat Keyword    10    Press Special Key    UP
    Press Special Key    DOWN
    Press Special Key    RIGHT
    Sleep    2
    Type With Modifiers    a    ALT
    Type With Modifiers    s    ALT
    Type With Modifiers    a    ALT
    Save Excel
    Excel.Application.Quit Application