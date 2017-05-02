# casestudy
Validations for security form using VBA
Following are the changes made to this application :-
Userform:--------Sedol : its max length property is set to 7Search Sedol : its max length property is set to 7Market cap: added validation to accept numeric values only, added in ValidateFrmSecurityContent 
Validation module-----------------Added validation logic for1) Price ccy : should select Price CCY from drop down only.2) Country of Issue : Should select country of issue from drop down only.
Created ValidateFrmSecurityContent function to check all validations at one go used this function in Save and Update button.
EraseData module----------------Created ClearFrmSecurityContents function to clear all the contents from Frmscecurity at one go.
