*** Settings ***
Library     ExcelLibrary
Library	    Collections
Library     OperatingSystem

*** Variables ***
${path_name}                C:/Users/theeradach_s/Desktop/python/ROBOT_FRAMEWORK/Excel - ReadWrite/Test Files/test_budget_import.xlsx
${sheet_name}               test data
${DEFAULT_SHEET_NAME}       SHEET_NAME_1

*** Test Cases ***
Check Correct Excel Doc
    ${document}=                    Create Excel Document    ${path_name}
    Should Be Equal As Strings      ${path_name}    ${document}
    Close All Excel Documents

Read Excel File
    Open Excel Document             ${path_name}            doc_id=docid
    ${rd}=	                        Read Excel Row	        row_num=5	max_num=3
    Log                             ${rd}
    ${cd}=                          Read Excel Column       col_num=2   max_num=3 
    Log                             ${cd}                   
    ${row_data}                     Create List             ${None}    12月    1月
    ${col_data}                     Create List             ${None}    関西    26期
    Lists Should Be Equal	        ${row_data}	${rd}		
    Lists Should Be Equal	        ${col_data}	${cd}	
    Close All Excel Documents	

Write Excel File 
    Create Excel Document	        doc_id=docname1		
    ${col_data}=	                Create List	a1	a2	a3
    Write Excel Column	            col_num=3	col_data=${col_data}
    Create Directory                Exported
    Save Excel Document	            filename=Exported/file.xlsx		
    Close All Excel Documents

Write Existed File                  # wirte file are not support uft-8  
    Open Excel Document             ${path_name}            doc_id=docid
    Write Excel Cell                col_num=4   row_num=1   value=Theeradach Subsin
    ${rc}                           Read Excel Cell         col_num=4   row_num=1
    Should Be Equal As Strings      ${rc}                   Theeradach Subsin
    Save Excel Document             filename=Exported/file2.xlsx	
    Close All Excel Documents