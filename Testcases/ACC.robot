*** Settings ***
Resource    ../Resources/Resources.robot
Variables    ../Resources/Locators.py

*** Variables ***
${Login_URL}    https://iea-nosso.sabacloud.com/Saba/Web_wdk/NA10P1PRD040/index/prelogin.rdf
${Org_URL}    https://iea-nosso.sabacloud.com/Saba/Web_spf/NA10P1PRD040/app/team/overview;spf-url=common%2Fteam%2Fteamhome%2Fxxemptyxx%2ForgAdmin%2Fbisut000000000003188
${Sheet_name}    ../Extracted Excel/ACC.xlsx
${start}     JOHN BUTCHER

*** Test Cases ***
Data from SabaCloud
    Extract Data From SabaCloud Broken    ${Org_URL}    ${sheet_name}    ${start}
#    Extract Data From SabaCloud    ${Org_URL}    ${sheet_name}