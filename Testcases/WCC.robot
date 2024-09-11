*** Settings ***
Resource    ../Resources/Resources.robot
Variables    ../Resources/Locators.py
Suite Teardown    Teardown words
Library    Collections

*** Variables ***
${Login_URL}    https://iea-nosso.sabacloud.com/Saba/Web_wdk/NA10P1PRD040/index/prelogin.rdf
${Org_URL}    https://iea-nosso.sabacloud.com/Saba/Web_spf/NA10P1PRD040/app/team/shared;spf-url=common%2Fteam%2Fteamhome%2Fxxemptyxx%2ForgAdmin%2Fbisut000000000003182
${Sheet_name}    ../Extracted Excel/WCC.xlsx
#    JOHN BUTCHER    # Set name here if you have to start the script using broken keyword.
${name_of_org}    WCC        # This will not required if you are not downloading attachments.
${MAX_RETRIES}    50          # Adjust according to need
${retry_count}    0          # Never Change this

*** Test Cases ***
Data from SabaCloud
    ${Employee_list} =    Create List
    Set Global Variable    ${Employee_list}
    ${Temp_Employee_list} =    Create List
    Set Global Variable    ${Temp_Employee_list}
    FOR    ${i}     IN RANGE    ${MAX_RETRIES}
        ${found} =    Set Variable    False
        Set Global Variable    ${found}
        ${length} =     Get Length    ${Temp_Employee_list}
        ${length} =     Evaluate    ${length} - 1
        Log    ${length}
        ${emp_name} =    Run Keyword If    ${length} > 0     Get From List    ${Temp_Employee_list}   ${length}
        TRY
            Log To Console    Starting the main loop for ${i}th time
            Close All Browsers
            Run Keyword If    '${i}' == '0'    Extract Data From SabaCloud    ${Org_URL}    ${sheet_name}    ${name_of_org}
            ${length} =     Get Length    ${Temp_Employee_list}
            ${length} =     Evaluate    ${length} - 1
            ${emp_name} =    Run Keyword If    ${length} > 0     Get From List    ${Temp_Employee_list}   ${length}
            Run Keyword If    '${i}' != '0'    Extract Data From SabaCloud Broken    ${Org_URL}    ${sheet_name}    ${emp_name}    ${name_of_org}
            BREAK
        EXCEPT
            ${retry_count}    Evaluate    ${retry_count}+1
            Set Global Variable    ${retry_count}
        END
        Log To Console    Re-running Testcase current retry count is ${retry_count}.
        Run Keyword If    '${retry_count}' == '${MAX_RETRIES}'    Fail
    END
    Log To Console    Script Execution Ended and ${retry_count} retries required.