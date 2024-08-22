*** Settings ***
Resource    ../Resources/Resources.robot
Variables    ../Resources/Locators.py

*** Variables ***
${Login_URL}    https://iea-nosso.sabacloud.com/Saba/Web_wdk/NA10P1PRD040/index/prelogin.rdf
${Org_URL}    https://iea-nosso.sabacloud.com/Saba/Web_spf/NA10P1PRD040/app/team/shared;spf-url=common%2Fteam%2Fteamhome%2Fxxemptyxx%2ForgAdmin%2Fbisut000000000003186
${Sheet_name}    ../Extracted Excel/IEAC1.xlsx
${start}    JAKE DAVIS

*** Test Cases ***
Data from SabaCloud
    Extract Data From SabaCloud Broken    ${Org_URL}    ${sheet_name}    ${start}