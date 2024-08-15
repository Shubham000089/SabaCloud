*** Settings ***
Resource    Resources.robot
Variables    Locators.py

*** Variables ***
${Login_URL}    https://iea-nosso.sabacloud.com/Saba/Web_wdk/NA10P1PRD040/index/prelogin.rdf

*** Test Cases ***
Data from SabaCloud
#    Landing on
    Extract Data From SabaCloud