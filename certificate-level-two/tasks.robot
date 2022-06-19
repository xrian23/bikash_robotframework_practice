*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    auto_close=${True}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF
#Library    RPA.Desktop
Library             RPA.Tables
Library             html_tables.py
Library             Collections
Library             DateTime
Library             OperatingSystem
Library             RPA.Archive
Library             RPA.Robocorp.Vault
Library             RPA.Dialogs

Suite Setup         Setup The Environment
Suite Teardown      Shut Down Browser


*** Variables ***
${url}                              ${EMPTY}
${ok_modal_button}                  xpath://button[normalize-space()='OK']
${order_csv_file}                   ${EMPTY}
${order_csv_file_download}          ${OUTPUT_DIR}${/}order.csv
${head_dropdownlist}                xpath://select[@id='head']
${button_hide_model_info}           xpath://button[normalize-space()='Hide model info']
${table_model_info}                 xpath://table[@id='model-info']
${input_legs_number_control}        xpath://input[@placeholder='Enter the part number for the legs']
${input_address_control}            xpath://input[@id='address']
${button_order_preview}             xpath://button[@id='preview']
${button_order}                     xpath://button[@id='order']
${button_order_another}             xpath://button[@id='order-another']
${html_locator_order_completion}    xpath://div[@id='order-completion']
${html_locator_robot_image}         xpath://div[@id='robot-preview-image']


*** Tasks ***
Order robots from RobotSpareBin Industries Inc
    Open the robot order website
    Close the annoying modal
    ${table_model_info}=    Read HTML table as RPA Table
    ${orders}=    Get orders
    ${str_curr_datetime}=    Get Current Date    time_zone=local    result_format=epoch    exclude_millis=${False}
    FOR    ${row}    IN    @{orders}
        Run Keyword And Continue On Failure    Fill the form    ${row}    ${table_model_info}
        ${is_visible}=    Is Element Visible    ${html_locator_order_completion}
        IF    ${is_visible} == ${True}
            ${pdf}=    Store the receipt as a PDF file    ${str_curr_datetime}${/}PDF${/}${row}[Order number]
            ${screenshot}=    Take a screenshot of the robot
            ...    ${str_curr_datetime}${/}SCREENSHOTS${/}${row}[Order number]
            Embed the robot screenshot to the receipt PDF file    ${pdf}    ${screenshot}
        END

        ${is_visible}=    Is Element Visible    ${button_order_another}
        IF    ${is_visible}==${True}
            Click Button When Visible    ${button_order_another}
        END
        ${is_visible}=    Is Element Visible    ${ok_modal_button}
        IF    ${is_visible}==${True}    Close the annoying modal
    END
    Create a ZIP file of the receipts    ${OUTPUT_DIR}${/}${str_curr_datetime}    ${str_curr_datetime}
    Log    Done.


*** Keywords ***
Open the robot order website
    Open Available Browser    ${url}    download=${True}
    #Open Browser    ${url}    browser=chrome    executable_path=${/}home${/}xrian${/}webdrivers${/}chromedriver
    Download    ${order_csv_file}    target_file=${order_csv_file_download}    overwrite=${True}

Get orders
    LOG    ${order_csv_file_download}
    ${table}=    Read table from CSV    header=${True}    path=${order_csv_file_download}
    Close Workbook
    RETURN    ${table}

Close the annoying modal
    Click Button When Visible    ${ok_modal_button}

Fill the form
    [Arguments]    ${row}    ${model_info}

    Click Element If Visible    ${head_dropdownlist}
    #Wait Until Element Is Visible    xpath=//option[normalize-space(va())='${row}[Head]']
    Select From List By Value    ${head_dropdownlist}    ${row}[Head]
    @{body_to_select}=    Find Table Rows    ${model_info}    1    ==    ${row}[Body]
    Log    @{body_to_select}
    ${body_number}=    Get From List    @{body_to_select}    1
    Select Radio Button    body    ${body_number}
    Input Text    ${input_legs_number_control}    ${row}[Legs]
    Input Text    ${input_address_control}    ${row}[Address]
    Click Button    ${button_order_preview}
    Click Button    ${button_order}

Store the receipt as a PDF file
    [Arguments]    ${pdf_filename_to_be_saved}
    Wait Until Element Is Visible    ${html_locator_order_completion}
    ${order_recdeipt_html}=    Get Element Attribute    ${html_locator_order_completion}    outerHTML
    Html To Pdf    ${order_recdeipt_html}    ${OUTPUT_DIR}${/}${pdf_filename_to_be_saved}.pdf
    RETURN    ${OUTPUT_DIR}${/}${pdf_filename_to_be_saved}.pdf

Take a screenshot of the robot
    [Arguments]    ${robot_image_to_be_saved}
    Wait Until Element Is Visible    ${html_locator_robot_image}
    ${robot_image_html}=    Get Element Attribute    ${html_locator_robot_image}    outerHTML
    Capture Element Screenshot    ${html_locator_robot_image}    ${OUTPUT_DIR}${/}${robot_image_to_be_saved}.png
    RETURN    ${OUTPUT_DIR}${/}${robot_image_to_be_saved}.png

Embed the robot screenshot to the receipt PDF file
    [Arguments]    ${pdf_order_receipt}    ${pdf_robot_image}
    ${files}=    Create List
    ...    ${pdf_robot_image}:x=0,y=0
    Add Files To Pdf    ${files}    ${pdf_order_receipt}    append=${True}

Get HTML TABLE MODEL info
    ${html_table}=    Get Element Attribute    xpath://table[@id='model-info']    outerHTML
    RETURN    ${html_table}

Read HTML table as RPA Table
    Click Button When Visible    xpath://button[normalize-space()='Show model info']
    ${html_table}=    Get HTML TABLE MODEL info
    ${table}=    Read Table From Html    ${html_table}
    ${dimensions}=    Get Table Dimensions    ${table}
    FOR    ${record}    IN    @{table}
        Log    ${record}
    END
    Click Button When Visible    xpath://button[normalize-space()='Hide model info']
    RETURN    ${table}

Create a ZIP file of the receipts
    [Arguments]    ${location_of_folder_to_zip}    ${zip_file_name}
    ${zip_file_name}=    Set Variable    ${OUTPUT_DIR}${/}${zip_file_name}.zip
    Archive Folder With Zip    ${location_of_folder_to_zip}    ${zip_file_name}    recursive=${TRUE}
    Remove Directory    ${location_of_folder_to_zip}    recursive=${True}

Shut Down Browser
    Close All Browsers

Setup The Environment
    Log    ${OUTPUT_DIR}
    ${secret}=    Get Secret    practice
    LOG    ${secret}[robotsparebinindustries_url]
    #Set Suite Variable    ${url}    ${secret}[robotsparebinindustries_url]
    Log    ${secret}[order_csv_file]
    Set Suite Variable    ${order_csv_file}    ${secret}[order_csv_file]
    Add text input    URL    label=Robot Order Poratl URL
    ${response}=    Run dialog
    LOG    ${response}
    Set Suite Variable    ${URL}    ${response.URL}

    #Remove Directory    ${OUTPUT_DIR}    recursive=${True}
