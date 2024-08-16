*** Settings ***
Resource    ../Resources/Resources.robot
Variables    ../Resources/Locators.py

*** Variables ***
${Login_URL}    https://iea-nosso.sabacloud.com/Saba/Web_wdk/NA10P1PRD040/index/prelogin.rdf
${Org_URL}    https://iea-nosso.sabacloud.com/Saba/Web_spf/NA10P1PRD040/app/team/shared;spf-url=common%2Fteam%2Fteamhome%2Fxxemptyxx%2ForgAdmin%2Fbisut000000000003240
${Sheet_name}    ../Extracted Excel/HBWHITE.xlsx

*** Test Cases ***
Data from SabaCloud
    Extract Data From SabaCloud    ${Org_URL}    ${sheet_name}