*** Settings ***
Documentation     Sikuli Library Demo
Library    SikuliLibrary
Library    String
Library    Collections
Library    Elements.py
Library    Process
Library    OperatingSystem
Library    pandas

*** Variables ***

*** Keywords ***


Open Planilha
    Run Process    python    excel.py
    
    

Coletar dados
    Run Process    python    Elements.py

Escrever dados
    Wait Until Screen Contain    export.png    15
    Click    export
    Paste Text    save.png    Planilha de teste
    Double Click    btn_save.png
    Press Special Key    ENTER
    Sleep    3
