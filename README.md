# PropertyManagement
A UI designed to provide ease of access regarding different properties stored in a database.

## Project Overview
PropertyManagement is primarily a VBA project that extracts property information from an excel sheet and stores relevent data in an Access database. Users interacting directly with the stored inofrmation in the spreadsheet was risky as changes were being made without protection or authorization. The purpose of this project was to limit any direct damage to this stored information by constructing a user interface that can be acessed by specific roles. After logging in, the UI allows the user to access property information for which they are authorized. Property information includes rent due, date of rent due, date of lease expiry, grace period, remaing balance and tenant information. In this way, the actual information from the ledger is no longer altered unless a role has permission to do so. There is also an added layer of privacy in regards to tenant data as that can also only be viewed or changed by specific roles. Not only does this make the data easier to read and understand, but it also limits users in what they can and cannot do. 

## Development Stack
![Access](https://img.shields.io/badge/access-%23A4373A.svg?style=for-the-badge&logo=microsoft-access&logoColor=white)


## Project Quickstart
Follow the steps below to run the project:
1. Clone the repository into a new folder on your desktop.
2. Open the file, "FinalProjectSER416_aarif6", and click the "Start Application" button on the sheet.
3. In the login box, use the following information to login as different roles:
   
     ```bash
     Role: Admin
     Username: Sparky
     Password: GoDevils
    ```
     ```bash
     Role: Suppport Staff
     Username: Gabe
     Password: thisIsFun

     or

     Username: Becca
     Password: reallyNeat
    ```
     ```bash
     Role: Analyst
     Username: Clarissa
     Password: iLoveExcel
    ```
     
5. Once login is complete, select a property from the list to view its data.

## Key Points
1. The Admin role has the most authority regarding the property management application. This means that the Admin is able to view each property listed in the database and alter or edit any information regarding the property or tenant as desired. This includes changing rent due, removing any balance, adding a new tennant to any property, altering the grace period atc. In addition, the Admin is also able to view and make changes directly to the spreedsheet whereas the other roles can only interact with the system through the user interface.
   
2. The role of the support staff is not as vast as that of the Admin. The staff can view and alter only the properties assigned to them. They can also make changes if desired and these changes are reflected in the spreadsheet and database. However, they cannot alter tenant information or view more secure information such as a tenant's SSN or private notes written by the Admin.
   
3. The Analyst's only role is to be able to view the generated reports regarding each property. The annual reports list total rent, total balance, total taxes and total additional charges for the whole year. The Admin has access to these reports as well. 

## Note:
Excel tends to block macros from running sometimes as they are considered a security risk for some of the older versions. If the macros are blocked when the file is downloaded, you can unblock them easily in a few steps:
1. Close the file, left click it and select "properties".
2. At the very bottom of the General tab in properties, there will be an "unblock" button. Clicking this will unblock the macros so the project can run smoothly.
3. After clciking "unblock", click "Apply" and save changes.
4. Reopen the excel file and click the "Start Application" button to begin.
