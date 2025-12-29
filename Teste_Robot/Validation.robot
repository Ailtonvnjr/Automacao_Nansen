*** Settings ***
Documentation     Sikuli Library Demo
Library    SikuliLibrary 




    


*** Variables ***

*** Keywords ***

    

Start here
    Click    start.png
    Wait Until Screen Contain   time.png    50
    ## Click    senha.png
    ### Input Text   senha.png    000000
    ## Press Special Key    ENTER



   
