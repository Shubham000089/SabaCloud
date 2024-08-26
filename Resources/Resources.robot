*** Settings ***
Library     SeleniumLibrary
Library  String
Library  Collections
Library    OperatingSystem
Variables    Locators.py
Library    Custom_Code.py

*** Variables ***
${Login_URL}    https://iea-nosso.sabacloud.com/Saba/Web_wdk/NA10P1PRD040/index/prelogin.rdf
${Username}    CERTTESTER
${Password}    P@ssw0rd123
${sr_no}    1
${found}    False

*** Keywords ***
# General
Open my browser
    [Arguments]    ${LOGIN_URL}
    # Local
#    Open Browser      ${LOGIN_URL}        headless Chrome
    Open Browser      ${LOGIN_URL}        Chrome
    Maximize Browser Window
    # Server
#    ${chrome_options} =     Evaluate    sys.modules['selenium.webdriver'].ChromeOptions()    sys, selenium.webdriver
#    Call Method    ${chrome_options}   add_argument    headless
#    Call Method    ${chrome_options}   add_argument    no-sandbox
#    Call Method    ${chrome_options}   add_argument    disable-dev-shm-usage
#    ${options}=     Call Method     ${chrome_options}    to_capabilities
#    Open Browser    ${LOGIN_URL}    browser=chrome    options=${chrome_options}
#    Maximize Browser Window
#    Set Window Size    ${1440}    ${900}

Add text in textbox
    [Arguments]    ${Locator}    ${Text}
    Wait Until Element Is Visible    ${Locator}
    Scroll Element Into View    ${Locator}
    Input Text    ${Locator}    ${Text}

Click Element After Visible
    [Documentation]    This will wait for 10 sec to visible the element after that click element.
    [Arguments]    ${Locator}
    Wait Until Page Contains Element    ${Locator}    timeout=60
    Wait Until Element Is Visible    ${Locator}    timeout=60
    Click Element    ${Locator}

# Specific
Extract Data From SabaCloud
    [Arguments]    ${Org_URL}    ${sheet_name}    ${name_of_org}
    Preloop process    ${Org_URL}    ${sheet_name}    ${name_of_org}
#    Loop of Organizations

Extract Data From SabaCloud Broken
    [Arguments]    ${Org_URL}    ${sheet_name}    ${start}    ${name_of_org}
    Preloop process broken    ${Org_URL}    ${sheet_name}    ${start}    ${name_of_org}

Preloop process
    [Arguments]    ${Org_URL}    ${sheet_name}    ${name_of_org}
    Open My Browser        ${Login_URL}
    Sign-In process
    Go To    ${Org_URL}
    Direct Link Loop    ${sheet_name}    ${name_of_org}

Preloop process broken
    [Arguments]    ${Org_URL}    ${sheet_name}    ${start}    ${name_of_org}
    Open My Browser        ${Login_URL}
    Sign-In process
    Go To    ${Org_URL}
    Direct Link Loop Broken    ${sheet_name}    ${start}    ${name_of_org}

Sign-In process
    Add text in textbox    ${TB_USERNAME}    ${Username}
    Add Text In Textbox    ${TB_PASSWORD}    ${Password}
    Click Element After Visible    ${B_SIGN_IN}

Home page to organization
    Click Element After Visible    ${B_MAIN_MENU}
    Press Keys    None    ARROW_DOWN    ARROW_DOWN    ENTER
    Waits for Stable Page of Organization
    Counting organization    

Waits for Stable Page of Organization
    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${LOADING}    timeout=5
    Wait Until Element Is Not Visible    ${LOADING}    timeout=120
    Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=120

Counting organization
    Open Organization Dropdown
    # Count Organisations
    Wait Until Element Is Visible    ${C_ORGANISATIONS}        timeout=10
    ${count_of_org} =     Get Element Count    ${C_ORGANISATIONS}
    Set Global Variable    ${count_of_org}

Open Organization Dropdown
    Wait Until Element Is Visible    ${IFRAME}    timeout=30
    Select Frame    ${IFRAME}
    Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=120
    Click Element After Visible    ${D_ORGANISATION_NAME}
    Click Element After Visible    ${DO_ORGANISATION}

Direct Link Loop
    [Arguments]    ${sheet_name}    ${name_of_org}
    [Documentation]    We are directly heading to link by avoiding organizational loop
    Waits for Stable Page of Organization
    Wait Until Element Is Visible    ${IFRAME}    timeout=30
    Select Frame    ${IFRAME}
    Get count of employee
    Getting data of all employee in list
    FOR    ${j}    IN    @{Employee_list}
        Append To List    ${Temp_Employee_list}        ${j}
        Log To Console    In Loop of employee for ${j}
        Waits for Stable Page of Organization
        # Check for Login page
        Unselect Frame
        ${check} =    Run Keyword And Return Status    Wait Until Element Is Visible    ${TB_USERNAME}
        Run Keyword If    "${check}" == "True"    Process of login to search box
        Run Keyword If    '${check}' == 'False'    Select Frame    ${IFRAME}
        # search
        Press Keys    ${EMP_SEARCH_BOX}    ${j}
        Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}   timeout=50
        Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=100
        Click Element After Visible        (//div[contains(@class, 'sd-team-user-detail')]//a)
        Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}    timeout=4
        Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=100
        ${check_completed} =    Run Keyword And Return Status     Wait Until Element Is Visible    ${C_COMPLETED}
        Run Keyword If    '${check_completed}' == 'True'    Further loop of certificates    ${sheet_name}    ${name_of_org}
        Run Keyword If    '${check_completed}' == 'False'   Exit From Employee
    END

Direct Link Loop Broken
    [Arguments]    ${sheet_name}    ${start}    ${name_of_org}
    [Documentation]    We are directly heading to link by avoiding organizational loop
    Waits for Stable Page of Organization
    Wait Until Element Is Visible    ${IFRAME}    timeout=30
    Select Frame    ${IFRAME}
    Get count of employee
    Getting data of all employee in list
    FOR    ${j}    IN    @{Employee_list}
#       For Running single broken list
        Log To Console    In Loop of employee for ${j}
        Run Keyword If    "${j}" == "${start}"    Loop
        Run Keyword If    '${found}' == 'True'    Temp adjustment    ${j}    ${name_of_org}
#       Ends here
    END

Loop of Organizations
    [Documentation]    Use if need to run through all the organizations.
#    FOR    ${i}    IN RANGE    1    ${count_of_org}+1
    FOR    ${i}    IN RANGE    3    5
        Log To Console    In Loop of organization for ${i} where total are ${count_of_org}
        Reload Page
        Waits for Stable Page of Organization
        Open Organization Dropdown
        Get the name of organization    ${i}
        Entry in Organization    ${i}
        Get count of employee
        Getting data of all employee in list
        # Start From here
        FOR    ${j}    IN    @{Employee_list}
#            # For Running single broken list
            Log To Console    In Loop of employee for ${j}
#            Run Keyword If    '${j}' == '${start}'    Loop
#            Run Keyword If    '${found}' == 'True'    Temp adjustment    ${i}    ${j}
#            # Ends here
            Waits for Stable Page of Organization
            # Check for Login page
            Unselect Frame
            ${check} =    Run Keyword And Return Status    Wait Until Element Is Visible    ${TB_USERNAME}    timeout=3
            Run Keyword If    '${check}' == 'True'    Process of login to search box    ${i}
            Run Keyword If    '${check}' == 'False'    Select Frame    ${IFRAME}
            # search
            Press Keys    ${EMP_SEARCH_BOX}    ${j}
            Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}   timeout=50
            Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=100
            Click Element After Visible        (//div[contains(@class, 'sd-team-user-detail')]//a)
            Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}    timeout=4
            Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=100
            ${check_completed} =    Run Keyword And Return Status     Wait Until Element Is Visible    ${C_COMPLETED}
            Run Keyword If    '${check_completed}' == 'True'    Further loop of certificates
            Run Keyword If    '${check_completed}' == 'False'   Exit From Employee
        END
        Exit From Organization
    END

Process of login to search box
    [Documentation]    If timed out get and came on login page, this process is followed
    Close Browser
    Open my browser    ${LOGIN_URL}
    Sign-In process
    Go to     ${Org_URL}
    Waits for Stable Page of Organization
    Wait Until Element Is Visible    ${IFRAME}    timeout=30
    Select Frame    ${IFRAME}

Temp adjustment
    [Arguments]   ${j}    ${name_of_org}
    Append To List    ${Temp_Employee_list}        ${j}
    Waits for Stable Page of Organization
    # Check for Login page
    Unselect Frame
#    Log To Console    Now Logout!!!!
#    Sleep    10
    ${check} =    Run Keyword And Return Status    Wait Until Element Is Visible    ${TB_USERNAME}
    Run Keyword If    '${check}' == 'True'    Process of login to search box
    Run Keyword If    '${check}' == 'False'    Select Frame    ${IFRAME}
    # search
    Press Keys    ${EMP_SEARCH_BOX}    ${j}
    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}   timeout=50
    Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=100
    Click Element After Visible        (//div[contains(@class, 'sd-team-user-detail')]//a)
    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}    timeout=4
    Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=100
    ${check_completed} =    Run Keyword And Return Status     Wait Until Element Is Visible    ${C_COMPLETED}
    Run Keyword If    '${check_completed}' == 'True'    Further loop of certificates    ${sheet_name}    ${name_of_org}
    Run Keyword If    '${check_completed}' == 'False'   Exit From Employee

Loop
    ${found} =    Set Variable    True
    Set Global Variable    ${found}

Getting data of all employee in list
    [Documentation]    Create list of all the employees
    ${count} =     Get Length    ${Employee_list}
    Run Keyword If    '${count}' == '0'    SubKeyword: Getting data of all employee in list

SubKeyword: Getting data of all employee in list
    FOR    ${l}    IN RANGE    1    ${integer_result}+2    # Count
        ${cnt_of_emp12} =     Get Text    ${PAGE_COUNT_EMP}
        ${range}    ${total}    Split String    ${cnt_of_emp12}    of
        ${start}    ${end}    Split String    ${range}    -
        ${start}    Strip String    ${start}
        ${end}    Strip String    ${end}
        Run Keyword And Ignore Error    Wait Until Element Is Visible    ${LOADING}    timeout=5
        Log To Console    Appending names in list ...
        Wait Until Element Is Not Visible    ${LOADING}    timeout=200
        ${Count_of_row} =     Get Element Count    (//div[contains(@class, 'sd-team-user-detail')]//a)
        FOR    ${m}    IN RANGE    1    ${Count_of_row}+1
            ${temp_name} =     Get Text    (//div[contains(@class, 'sd-team-user-detail')]//a) [${m}]
            Append To List    ${Employee_list}        ${temp_name}
        END
        Exit For Loop If    ${Count_of_row} < 10
        ${check} =     Run Keyword And Return Status    Should Be True    ${start} <= ${x} <= ${end}
        Unselect Frame
        Select Frame    ${IFRAME}
        Run Keyword If    '${check}' == 'False'    Click Element    ${B_NEXT}
    END
    Set Global Variable    ${Employee_list}
    ${length_of_list} =     Get Length    ${Employee_list}
    Log To Console    List Created !!!

Further loop of certificates
    [Arguments]    ${sheet_name}    ${name_of_org}
    [Documentation]    When u get land in employee, this process starts
    Get employee name and person no.
    Preloop for certification
    Check Load More Condition
    Press Keys    None    PAGE_DOWN
    ${cnt_of_view} =     Set Variable    ${1}
    Set Global Variable    ${cnt_of_view}
    ${check1} =     Run Keyword And Return Status    Wait Until Element Is Visible    //span[contains(text(), 'No data available for the selected filter.')]    timeout=3
    Run Keyword If    '${check1}' == 'True'    Click Element    (//div[contains(@id, 'metafilter')]//div[contains(@id, 'toolbar')]//span[contains(text(), 'Apply filter')]//..) [1]
    FOR    ${k}    IN RANGE    1    ${c_completed_1}+1
#    FOR    ${k}    IN RANGE    1    2
        Log To Console    In Loop of certificates for ${k} where total are ${c_completed_1} and count of view is ${cnt_of_view}
        ${check} =    Run Keyword And Return Status    Wait Until Element Is Visible    (//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]    timeout=5
        Run Keyword If    '${check}' == 'False'    Press Keys    None    PAGE_DOWN
        ${check1} =     Run Keyword And Return Status    Wait Until Element Is Visible    ${LOAD_MORE}    timeout=2
        Run Keyword If    '${check1}' == 'True'    Action on load more
        Wait Until Page Contains Element    (//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]    timeout=40
        Wait Until Element Is Visible    (//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]    timeout=40
        ${check} =     Run Keyword And Return Status    Click Element    (//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]
        Run Keyword If    '${check}' == 'False'    Wait Until Element Is Not Visible    ${CERTIFICATE_TOOLTIP}    timeout=30
        Run Keyword If    '${check}' == 'False'    Click Element    (//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]
        Get Certificate Name    ${k}
        Press Keys    None    TAB
        Sleep    0.3
        Press Keys    None    TAB
        Sleep    0.3
        Press Keys    None    ARROW_DOWN
        Sleep    0.3
        Check View is available    ${k}    ${sheet_name}    ${name_of_org}
    END
    Exit From Employee

Check Load More Condition
    [Documentation]    Check load more available on screen
    ${check} =     Run Keyword And Return Status    Wait Until Element Is Visible    ${LOAD_MORE}    timeout=10
    Run Keyword If    '${check}' == 'True'    Action on load more

Action on load more
    [Documentation]    If found load more takes further action
    FOR    ${i}    IN RANGE    1    500
        Click Element    ${LOAD_MORE}
        ${check1} =     Run Keyword And Return Status    Wait Until Element Is Visible    ${LOAD_MORE}    timeout=10
        Exit For Loop If    '${check1}' == 'False'
    END
        
Get the name of organization
    [Arguments]    ${i}
    Wait Until Element Is Visible   (//div[@class='sd-team-org-inner-container']//a) [${i}]//span    timeout=20
    ${name_of_org} =     Get Text   (//div[@class='sd-team-org-inner-container']//a) [${i}]//span    # 1
    Log To Console    ${name_of_org}
    Set Global Variable    ${name_of_org}

Entry in Organization
    [Arguments]    ${i}
    Click Element After Visible    (//div[@class='sd-team-org-inner-container']//a) [${i}]
    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${PLEASE_WAIT}
    Run Keyword And Ignore Error    Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=140

Get count of employee
    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${LOADING}    timeout=3
    Wait Until Element Is Not Visible    ${LOADING}    timeout=120
    Wait Until Element Is Not Visible    ${PLEASE_WAIT}    timeout=120
    Wait Until Element Is Visible    ${PAGE_COUNT_EMP}    timeout=50
    ${cnt_of_emp1} =     Get Text    ${PAGE_COUNT_EMP}
    ${parts} =    Split String    ${cnt_of_emp1}    of
    ${pg_num_emp} =    Strip String    ${parts}[1]
    Set Global Variable    ${pg_num_emp}
    ${result}    Evaluate    ${pg_num_emp} / ${10}
    ${integer_result}    Evaluate    int(${result})
    Set Global Variable    ${integer_result}

Entry in employee detail page
    Click Element After Visible        (//div[contains(@class, 'sd-team-user-detail')]//a) [1]

Get employee name and person no.
    Waits for Stable Page of Organization
    Wait Until Element Is Visible    (//h2) [1]        timeout=30
    ${emp_url} =    Get Location
    Set Global Variable    ${emp_url}
    Click Element After Visible    ${L_PROFILE}
    Run Keyword And Ignore Error    Wait Until Element Is Visible    ${LOADING}    timeout=2
    Wait Until Element Is Not Visible    ${LOADING}    timeout=150
    
    Wait Until Element Is Visible    (//h2) [1]        timeout=30
    ${emp_name} =     Get Text    (//h2) [1]
    Set Global Variable    ${emp_name}
    
    Click Element After Visible    ${MORE_INFO}
    Wait Until Element Is Visible    ${C_PERSON_NO}

    ${text_person_no} =     Get Text    ${C_PERSON_NO}
    Set Global Variable    ${text_person_no}
    Click Element After Visible    ${B_CLOSE}
    ${check} =     Run Keyword And Return Status    Wait Until Element Is Not Visible    ${B_CLOSE}
    Run Keyword If    '${check}' == 'False'    Click Element After Visible    ${B_CLOSE}

Exit from employee
    ${check} =     Run Keyword And Return Status    Click Element After Visible    ${B_BACK}
    Run Keyword If    '${check}' == 'False'    Press Keys    None    PAGE_UP
    Run Keyword If    '${check}' == 'False'    Click Element After Visible    ${B_BACK}

Get name of employee
    [Arguments]    ${j}
    Wait Until Element Is Visible    (//div[contains(@class, 'sd-team-user-detail')]//a) [${j}]
    ${emp_name} =     Get Text    (//div[contains(@class, 'sd-team-user-detail')]//a) [${j}]
    Set Global Variable    ${emp_name}

Preloop for certification
    Click Element After Visible    ${L_PLAN} 
    Unselect Frame
    Wait Until Element Is Not Visible    ${LOADING}    timeout=200
    Wait Until Element Is Visible    ${IFRAME}    timeout=20
    Select Frame    ${IFRAME}
    Sleep    0.5
    ${c_completed_1} =     Get Text    ${C_COMPLETED}
    Set Global Variable    ${c_completed_1}

    Press Keys    None    PAGE_DOWN
    Sleep    2
    Click Element After Visible    ${COMPLETED}
    ${check1} =    Run Keyword And Return Status    Wait Until Element Is Visible    ${LABEL_TYPE}    timeout=10
    Run Keyword If    '${check1}' == 'False'    Click Element    ${COMPLETED}

    ${check} =    Run Keyword And Return Status    Wait Until Element Is Visible    ${TEXT_NO_DATA}    timeout=1.5
    Run Keyword If    '${check}' == 'True'    Click Element    ${B_APPLY_FILTER}

Get Certificate Name
    [Arguments]    ${k}
    [Documentation]    Get the certificate name
    Wait Until Element Is Visible    ((//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]/ancestor::tr[not(.//a[contains(@class, 'goal')])]//b) [1] | ((//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]/ancestor::tr[.//a[contains(@class, 'goal')]]//a[contains(@class, 'goal')]) [1]    timeout=30
    ${c_certificate_name} =     Get Text     ((//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]/ancestor::tr[not(.//a[contains(@class, 'goal')])]//b) [1] | ((//table[contains(@id,'sjsgridview')]//tr//span[contains(@class,'pendingApprovalIcon')]) [${k}]/ancestor::tr[.//a[contains(@class, 'goal')]]//a[contains(@class, 'goal')]) [1]
    Set Global Variable    ${c_certificate_name}

Check View is available
    [Arguments]    ${k}    ${sheet_name}    ${name_of_org}
    Log    ${cnt_of_view}
    ${check} =     Run Keyword And Return Status    Wait Until Element Is Visible    (//div[contains(@id, 'sjsmenu')]//a[contains(@id, 'menuitem')]//span[normalize-space(text())='View']) [${cnt_of_view}]    timeout=1.5
    Run Keyword If    '${check}' == 'True'    Check Attachment available    ${sheet_name}    ${name_of_org}
    Run Keyword If    '${check}' == 'False'    Move to next certificate    ${sheet_name}

Check Attachment available
    [Arguments]    ${sheet_name}    ${name_of_org}
    Click Element    (//div[contains(@id, 'sjsmenu')]//a[contains(@id, 'menuitem')]//span[normalize-space(text())='View']) [${cnt_of_view}]

    # Increase Count By 1
    ${cnt_of_view} =     Evaluate    ${cnt_of_view} + 1
    Set Global Variable    ${cnt_of_view}

    # Check for Attachment
    ${check_attachment} =     Run Keyword And Return Status    Wait Until Element Is Visible    ${ATTACHMENT}    timeout=2
    Run Keyword If    '${check_attachment}' == 'True'    Action on Attachment to download certificates    ${sheet_name}    ${name_of_org}
    Run Keyword If    '${check_attachment}' == 'False'    Write Data In Excel At Attachment    ${sheet_name}

    ${check} =     Run Keyword And Return Status    Wait Until Element Is Visible    ${B_MODULE_CLOSE}
    Run Keyword If    '${check}' == 'True'    Click Element    ${B_MODULE_CLOSE}
    
    ${check1} =     Run Keyword And Return Status    Wait Until Element Is Not Visible    ${B_MODULE_CLOSE}    timeout=1.5
    Run Keyword If    '${check1}' == 'False'    Click Element    ${B_MODULE_CLOSE}

Action on Attachment
    [Arguments]    ${sheet_name}
    ${attachment_name} =     Get Text    ${ATTACHMENT}
    Set Global Variable    ${attachment_name}
    Write Data In Excel At Completion    ${sheet_name}

Action on Attachment to download certificates
    [Arguments]    ${sheet_name}    ${name_of_org}
    ${attachment_name} =     Get Text    ${ATTACHMENT}
    Set Global Variable    ${attachment_name}
    Click Element    ${ATTACHMENT}
    Sleep    2
    Switch Window    new

# change the folder name on 424, 425, 428 when you run next company

    ${Check_path} =    Run Keyword And Return Status     Directory Should Exist    C:\\SabaCloud_Reports\\${name_of_org}
    Run Keyword If    '${Check_path}' == 'False'    Create Directory    C:\\SabaCloud_Reports\\${name_of_org}
    Run    python -c "import pyautogui; pyautogui.hotkey('ctrl', 's')"
    ${random} =    Custom_Code.gen_random_string
    ${resolved_path}=    Set Variable    C:\\SabaCloud_Reports\\${name_of_org}\\${random} ${attachment_name}
    Sleep    1.5
    Run    python -c "import pyautogui; pyautogui.typewrite('${resolved_path}')"
    Sleep    1
    FOR    ${i}    IN RANGE    1    3
        Run    python -c "import pyautogui; pyautogui.press('tab')"
        Sleep    0.5
    END
    Run    python -c "import pyautogui; pyautogui.press('enter')"
    Sleep    1.5

    Write Data In Excel At Completion    ${sheet_name}
    Close Window
    switch window       title=Blueprint
    Wait Until Element Is Visible    ${IFRAME}
    Select Frame    ${IFRAME}
    Sleep    0.5

Move to next certificate
    [Arguments]    ${sheet_name}
    Press Keys    None    ARROW_LEFT
    Sleep    0.5
    Write Data In Excel At View    ${sheet_name}

Exit From Organization
    Unselect Frame
    Click Element After Visible    ${B_MAIN_MENU}
    Press Keys    None    ENTER
    Waits for Stable Page of Organization

Write Data In Excel At View
    [Arguments]    ${sheet_name}
    Common Excel Data    ${sheet_name}
    Custom_Code.add_value    G    No    ${sheet_name}
    Custom_Code.add_value    H    No    ${sheet_name}
    Custom_Code.add_value    I    None    ${sheet_name}
    ${sr_no} =     Evaluate    ${sr_no} + 1
    Set Global Variable    ${sr_no}

Write Data In Excel At Attachment
    [Arguments]    ${sheet_name}
    Common Excel Data    ${sheet_name}
    Custom_Code.add_value    G    Yes    ${sheet_name}
    Custom_Code.add_value    H    No    ${sheet_name}
    Custom_Code.add_value    I    None    ${sheet_name}
    ${sr_no} =     Evaluate    ${sr_no} + 1
    Set Global Variable    ${sr_no}

Write Data In Excel At Completion
    [Arguments]    ${sheet_name}
    Common Excel Data    ${sheet_name}
    Custom_Code.add_value    G    Yes    ${sheet_name}
    Custom_Code.add_value    H    Yes    ${sheet_name}
    Custom_Code.add_value    I    ${attachment_name}    ${sheet_name}
    ${sr_no} =     Evaluate    ${sr_no} + 1
    Set Global Variable    ${sr_no}

Common Excel Data
    [Arguments]    ${sheet_name}
    Custom_Code.add_value    A    ${sr_no}   ${sheet_name} 
    ${timestamp} =     Custom_Code.get_current_date_time
    Custom_Code.add_value    B    ${timestamp}    ${sheet_name}
    Custom_Code.add_value    C    ${emp_url}    ${sheet_name}
    Custom_Code.add_value    D    ${emp_name}    ${sheet_name}
    Custom_Code.add_value    E    ${text_person_no}    ${sheet_name}
    Custom_Code.add_value    F    ${c_certificate_name}    ${sheet_name}

Teardown words
    Run Keyword If All Tests Passed    Log To Console       Testcase is passed successfully
    Run Keyword If Any Tests Failed      Log To Console    All retries failed after ${retry_count} attempts.