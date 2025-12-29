*** Settings ***
Resource    resource/Scenaries.resource
Resource    Check.robot




*** Test Cases ***
Caso de teste 1 - Validation
    [Documentation]  Este testes visa verificar as peças que flharam
    [Tags]           Login
    Falhado
    Capturar ID
    