# database-integration

1. Please remove  " , \ and #Value! From the excel file, delete all no value columns and make sure each column is a value but not a equation.

2. In the program, you can manually enter the corresponding head of each column. If there is no input, the system will try to find the keywords in the first row and try to distinguish the heads. If none of them works, program will terminate.

3. In the program, you can manually set the default value for each column. Otherwise the default value will all be None(NULL). The default first data value line is second row which can be set to any valuable row.

4. When dealing with the company name and entity name, company name end with "pty ltd" or "p/l" will create a entity name in Entity table, no matter it is really a entity name or not. Then "pty ltd" will be remove from name to save as a business name in Company table.


Company Table Only

5. When dealing with company address, the provided company name address will fist be send to google map api for check. If google map api cannot find that address, the company name will be send to google location api for check.
If both apis do not work, the data will be stored with its original address or with postcode "0000" if none address is provided. These unrecognized data will be recorded and put into an excel file "fail list [time].xlsx" for manual inspection later. 
Program won't update address information for those data with postcode "0000" as they are unrecognized. When a new data is read which contain address, a new recoed will be created but not update that "0000" one. Please delete them or update them manually.

6. Companies are distinguished by company_name and postcode. Branches of the same company in different subs will be recorded. 

7. If mutiple phone, fax numbers or email addresses are provided, only one of them will be recorded, separated by ","  ";"  or "/". 

8. Comment will be store in the form: "[comment1 head] : [comment1 content] --- [comment2 head] : [comment2 content] --- ".


People Table only

9. People are distinguished by full_name and entity_id. People with same name work for different company will be stored. If no company name is provided, the data will still be stored, but have to be updated manually later. Program won't update company information for them as they are not recognized.

10. State of people are not a column of People table, but you can find them in comment.

11. If excel provides full name, family name and last name but full name != family name + last name, the record will be stored based on full name. Family name and last name will be replace with full name's split. If only family name or only last name is provided, the data won't be stored as this guy doesn't have a full name.