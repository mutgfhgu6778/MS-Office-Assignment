# 留学生日常作业/课程设计/代码辅导/CS/DS/商科，仅为漂洋过海的你❥
标签：Computer Science、Data Science、Business Administration，留学生编程作业代写&&辅导

## 个人介绍:
本人是一名资深码农，酷爱投资。

[为您提供 CS , DS , 商科 编程作业代写](http://dzuoye.work "编程代写")

<img src="design2023866.jpg"  width="200" />
# Background and Scenario 
An innovative, fairly new, company called FreshHarvest Grocers (FHG) has been set up that 
provides a door-to-door grocery delivery service to the inner-city suburbs of Brisbane. FHG is 
set up as a franchise, which means that there is a ‘head’ franchisor with several franchisees who 
operate ‘under license’ from the franchisor. FHG has four franchisees in Brisbane. Each 
franchisee is a local grocery shop. 
As franchisees of FHG, they are licenced to deliver daily groceries by the box to homes and 
businesses in nearby suburbs (their ‘franchise area’). Over its two years of operation, FHG has 
built up a trustworthy reputation among its customer bases. FHG customers receive a ‘set 
grocery’ box of groceries each week via their membership subscription program. The length of 
membership varies in terms of 14, 21, or 42 weeks, and subscription fees vary by membership 
terms. 
Harvey Thompson, the owner of FHG, has asked you to develop a spreadsheet that will help 
refine the franchise area and lower the distance travelled. Harvey is very environmentally 
conscious and does not want to damage the planet to deliver groceries. He wants you to: 
(1) Develop a schedule of employee budgeted salary costs according to his specifications 
and build a summary table using database functions; 
(2) Undertake a Solver analysis on the business franchise areas to determine a reallocation 
of franchise areas by distance from the store; 
(3) Undertake a scenario analysis on saving monthly/fortnightly for building a new store 
that will cost a ‘What If’ certain amount of money in a few years; 
(4) Provide some business-focused comments to Harvey relating to this MS Excel solution.

# List of Sheets in Excel Workbook 
When submitted, your final solution will have the following sheets: 
• Document Control 
• Constant 
• Employees 
• Payroll Summary 
• Pivot Table 
• Pivot Chart 
• Current Franchise Distribution 
• Franchise Redistribution 
• Solver Analysis Answer Report 
• Pivot Table 
• Pivot Chart 
• New Store Investment 
• Scenario Summary 
• Comments to Harvey 
Sheets in italics need to be created by following instructions in this document, as they are not 
in the template file.
# Document Control Sheet 
Hint: Throughout the spreadsheet, cells with a light shaded blue background require you 
to enter a value or a formula in them or take some actions with them. 
Cells with a light-orange background are to be populated by either the Solver or Scenario 
Manager tools, not you. 
Cells with no colour background should not be edited or changed by you unless explicitly 
directed to do so in this specification document. 
First enter your details: Student name and student number. 
In addition, you should list any assumptions that you have made when you developed your 
assignment on this sheet. The assumptions allow examiners to understand your work in context. 
You should use these assumptions to resolve any ambiguities you might identify in this Case 
Specification. 
The assumptions you make must be logical and consistent with the scenario provided in this 
Case Specification. 
If you do not make any assumptions, please leave this section empty.

# Constant Sheet 
This sheet contains all the lookup tables that you will need to use in the assignment. When 
using lookup tables in your formulas from the Constant sheet, make sure they are accessed 
using appropriate Named Ranges. 
You should also edit some tables which are empty and coloured in blue and then format this 
sheet professionally. 
Note: Throughout this assignment, you must use a Named Range whenever referring to 
a range or a cell in writing formulas/functions to ensure that whoever reviews your 
spreadsheet can understand it, especially when the range or cell is from another 
worksheet. You can also use Named Ranges to refer to a range or a cell in the same 
worksheet, but this is optional. 
MARKS ARE ALLOCATED IN THE MARKING RUBRIC FOR USING NAMED 
RANGES TO REFER TO RANGES OR CELLS IN A DIFFERENT WORKSHEET 
There are 8 lookup tables or values contained in this Constant Sheet. You are to complete these 
as directed below.
# Employee Salary Table 
Employees are paid at different rates based on their job title. Each job comes with a different 
employer superannuation percentage rate and different email domain. The details of the 
different job descriptions are presented below. All data is fictious in the tables below.
You are required to complete the data entry of the table in the workbook.
Table 1: Employee Salary Table for 2022-23

|Job Title | Annual Salary| Employer Super| Department |
|------|------|------|------|
|Accountant| $63,753 |10.5% |Accounting|
|Operation Manager |$66,824 |12% |Operation|
|Owner |$121,765 |17% |Executives|
|Delivery Service Manager |$68,681 |11.5% |Operation|
|IT Manager |$75,389 |13.5% |IT|
|Franchisee Manager |$98,985 |17% |Operation|
|Senior Delivery Service Manager |$76,789 |13.5% |Operation|

# Annual Tax Table 
Tax is withheld using the following tax rates for 2022-23. This information has been entered 
for you in the Constants Sheet. 
Table 2: Australian Taxable Income Table for 2022-23

|Taxable Income |Tax on this Income|
|------|------|
|$0 - $10,285 |Nil| 
|$10,286 - $41,785 |12c for each $1 over $10,285|
|$41,786 - $89,085 |5,092 plus 22c for each $1 over $41,785|
|$89,086 - $170,050 |29,467 plus 24c for each $1 over $89,085|

# Employee Superannuation Contribution Table 
Employees at FHG have collectively agreed to contribute a percentage of their annual salary to 
their superannuation fund based on their age at the beginning of the financial year as a posttax contribution (‘non-concessional contributions’). In Australia, a superannuation fund 
receives a percentage of every salary to employees to invest on their behalf so that they can 
draw on it when they retire.
You are required to complete the data entry of the table in the workbook.
o Employees aged 30 and over have elected to contribute 4.0%. 
o Employees aged 40 and over have elected to contribute 4.5%. 
o Employees aged 50 and over have elected to contribute 5.0%. 
o Employees aged 60 and over have elected to contribute 5.5%. 
Note: In Australia, the financial year is for the period 1 July to 30 June, which is different to 
the calendar year, which is for the period 1 January to 31 December. 

# Christmas Bonus Rates Table 
Employees at FHG who have had extended service with the company are paid an annual 
Christmas bonus at the end of each calendar year. You are required to complete the data 
entry of the table in the workbook. 
o Employees who have been employed for at least 3 years at the beginning of the calendar 
year receive a 4% bonus on their annual salary. 
o Employees who have been employed for at least 5 years at the beginning of the calendar 
year receive a 5% bonus on their annual salary. 
o Employees who have been employed for at least 7 years at the beginning of the calendar 
year receive a 6% bonus on their annual salary. 
o Employees who have been employed for at least 9 years at the beginning of the calendar 
year receive a 7% bonus on their annual salary. 
o Employees who have been employed for at least 11 years at the beginning of the 
calendar year receive a 8% bonus on their annual salary.

# Beginning of Calendar Year 
Use function to enter the first day of the 2023 calendar year (i.e., 01/01/2023). 
You are required to complete the data entry of the table in the workbook. 
Beginning of Financial Year 
Use function to enter the first day of the 2023/2024 financial year (i.e., 01/07/2023). 
You are required to complete the data entry of the table in the workbook. 
# FHG Subscriptions Table 
The subscription fee paid by customers varies according to the number of weeks they subscribe. 
Customers pay $90 per week for a 14-week subscription, $75 per week for a 21-week 
subscription, and $50 per week for a 42-week subscription. This information has been entered 
for you in the Constants Sheet. 

# Distance Survey and Suburb Profile Table 
Previously, franchise areas were allocated according to a rule of thumb (‘whatever worked’) at 
the time the franchise was allocated. As FHG matures, Harvey now wants to consider allocating 
franchise areas based on the average actual travel distance from the shop to the suburbs that 
they service. 
This table is central to those calculations. 
Each row in this table is an inner-city suburb in Brisbane that is within 6kms or so of the 
Brisbane CBD. The latitude and longitude of an ‘average’ (centroid) point for each suburb is 
provided. You are to use this information to determine distance for franchise areas. 
Each row also indicates the prospective subscribers to the FHG service in these Brisbane 
suburbs to each subscription type (14, 21, or 42 weeks). This information is derived from 
extensive and, according to Harvey, infallible, market research 1 . Prospective subscribers
contain the number of potential subscribers in each suburb, as discovered through market 
research. The role of prospective subscribers versus actual subscribers is discussed below in 
the Current Franchise Distribution section. 
Each column in this table represents the four (4) current franchisee stores in Brisbane (Brisbane 
City, South Brisbane, Milton, and Fortitude Valley). 
In this table, you are to calculate the distance from each franchisee store to each suburb using 
the latitude and longitude. To do this, use the latitude and longitude of each location according 
to the following formula: 
𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷𝐷 = ACOS(COS( RADIANS(90 − 𝐿𝐿𝐿𝐿𝐿𝐿1) ) × COS(RADIANS(90 − 𝐿𝐿𝐿𝐿𝐿𝐿2))
+ SIN(RADIANS(90 − 𝐿𝐿𝐿𝐿𝐿𝐿1))
× SIN(RADIANS(90 − 𝐿𝐿𝐿𝐿𝐿𝐿2)) × COS(RADIANS(𝐿𝐿𝐿𝐿𝐿𝐿𝐿𝐿1 − 𝐿𝐿𝐿𝐿𝐿𝐿𝐿𝐿2)))
× 6371
This Haversine formula uses a widely accepted method of calculating distances across the 
surface of the earth (which looks quite complicated because the earth is not flat). Do not worry 
about the calculation method. Just insert the latitude and longitude of each suburb and each 
franchisee store respectively into your formula as follows.
Note: The Shop Code at the top of this table relates each Shop Code to the suburb in 
which it is located. 
 
Hint: You may wish to check your calculations of distance between suburbs using Google 
Earth to measure point-to-point distance. Google Maps would provide travel distance 
using roads, which would be longer than the Haversine formula (great circle method) 
above. 
 
Note: The distance between a franchisee store in a suburb and the suburb in which it is 
located will be 0. For example, the distance between the suburb of Milton in the row of 
this matrix and the suburb of Milton in the column of this matrix should be 0. If it is not 
zero – then your formula is definitely wrong! 
 
Note: Format this table appropriately. 

# Employees Sheet 
Note: Please note that your formulas should be efficient and not hardcoded. You can use 
a Lookup and Reference function to achieve this. 
When using lookup tables in your formulas for values in the Employees Sheet, make sure 
they are accessed using appropriate named ranges. 
You should also format this sheet professionally and meaningfully. 
The employee sheet keeps track of FHG’s employees. Your tasks are to: 
Suggested order to complete tasks: Employee Budgeted Salary Costs -> Employee 
Personal Details -> Database Functions 
Employee Budgeted Salary Costs 
(1) Column M: Insert a formula to retrieve the annual salary of the employee from the 
Constant Sheet based on employee’s Job Title. 
 
(2) Column N & O: Following this, insert formulas to calculate the employer and employee 
superannuation contributions. 
Note: Please note that employer super contribution rates are in the salary table. The 
employer superannuation contribution rates must be multiplied by annual salary. 
Hint: You will need to use the employee’s birthday in relation to the first day of the 
financial year to calculate their employee superannuation contributions. The employee 
superannuation contribution is post-tax contribution. Employees younger than 30 would 
not pay any extra contribution.
Take a look at the date and time functions to best understand how to calculate accurate 
durations between dates, dividing number of days by 365 is not accurate enough.
(3) Column P: Insert a formula to determine the Christmas bonus employees receive in 
addition to their annual salary. 
Hint: You will need to use the employee’s first working day in this formula in relation to 
the beginning of Calendar Year. Take a look at the date and time functions. 
(4) Column Q: Using a formula calculate the annual income tax & Medicare levy withheld 
from employees based on their salary. 
Note: In your solution, assume that all employees pay the Medicare levy of 3% (that is, 
assume all employees earn more than the threshold for low-income earners, and no 
Medicare Levy Surcharge applies). The taxable income should include Christmas bonus.
 
Hint: Use the Annual Tax Table to calculate Income Tax from all income figures. For 
example, an employee whose salary is $68,720 who has been at the store for three years 
would receive ($68,720+ (0.04 x $68,720)) = $71,468.8 in taxable income. 
On this taxable income, the accountant would pay income tax of ($5,092 + ($71,468.8 – 
41,785) x 0.22) = $11,622.44. The Medicare Levy of 3% also applies and so Income Tax 
& Medicare Levy would be $11,622.44 + ($71,468.8 x 0.03) = $13,766.504. 

(5) Column R: Finally, insert a formula to determine the annual 
take home balance for each employee – this is each employee’s total income less tax 
paid less any employee contributions to superannuation. 
# Employee Personal Details 
(1) Email Address (Column K): Insert a formula to form out email address for each 
employee. All employees share the same company domain, which is “fhg.com.au”, but 
different subdomains for each department. 
Hint: Derive this column from ‘Given Name’ (Column C), ‘Surname’ (Column E), and 
‘Employee Salary Table’ on the Constant sheet. You may need to look up the department 
name and then connect with the company domain. For example, the email address for 
employee Eleanor Greene who serves as a delivery service manager should be 
E.Greene@operation.fhg.com.au
(2) Residential Suburb (Column L): Insert a formula to determine the employee’s 
residential suburb (i.e., Milton). 
Hint: You will need to search for the suburb name in the employee’s address to determine 
the suburb of their home address. Here you are allowed to explore some Text Functions 
which are not covered in tutorial.
# Database Functions 
You then need to complete the FHG Summary Table at the top of the sheet using Database 
Functions. Please use the criteria range (A2:R3) and database range (A17:R94) by assigning 
the named ranges “Criteria” and “Database” respectively, apply database functions to 
calculate the minimum, maximum, average, and total values for the listed headings (MR), so that changing the criteria in row 3 results in changes to the summary table values in rows 
9-12. You should also use a database function to count the number of records in the 
schedule that meet the criteria for cell M7. 
Hint: Database functions begin with a “D” and rely on criteria set out in a range that you 
identify and populate with your criteria. You will need to apply the advanced filter 
function in order to filter out records that meet the criteria.
The formulas should be robust to control error with IFERROR. 
When submitting the spreadsheet, set the criteria for summary table database functions so 
that the Summary Table displays data relating to only the Job Title of “Delivery Service 
Manager”. 
# Payroll Summary Sheet 
Payroll Summary sheet records the actual payment across 12 months for each employee in 2023 
(Please note that the calculations in Employees Sheet are budgeted take-home payment not the 
actual payment employees received)
Using the information from the Employees Sheet and the Payroll Summary Sheet, create 
Relationship through Power Pivot and generate a PivotTable to compare the budgeted takehome payment versus actual payment based on the Job Title. Calculated fields in PivotTable 
are not required but please make sure the PivotTable has meaningful label headers. To make 
it easier for other users to view the movement of the salaries, visualise the PivotTable by 
inserting a Clustered Column PivotChart. The PivotChart must have a meaningful chart title. 
You should put the Pivot Table and the Pivot Chart in two separate worksheets and give each 
of the two worksheets a meaningful sheet name. 
Hint: You need to import the data from Payroll Summary 2023 Sheet. Please clean the 
dataset before loading into the sheet, removing duplicates. PivotTable and Pivot Chart
worksheets do not exist in the template. ‘Job Title’ should be displayed in Rows. 
 
Hint: You do not need to use calculated fields in the PivotTable, which means that you 
do no need to add columns/rows manually in PivotTable.
# Current Franchise Distribution Sheet 
Currently, FHG stores are assigned suburbs as their franchise area (where they have exclusive 
rights to provide FHG services) in an ad hoc manner. Harvey does not like the fairly random 
manner by which this allocation was made. 
You are to model the Current Franchise Distribution and calculate the ‘Total Number of 
Customers’, ‘Total Distance’, and ‘Total Subscription Revenue’ for each store using the 
layout in this sheet. In doing so, calculate the ‘Distance’, Number of actual subscribers by 
each subscription type, and ‘Revenue’ in this sheet for all suburbs of FHG. 
Hint: You will need to use some of Excel’s Lookup & Reference, Math & Trig, and 
Logical functions in some of your formulas. 
Hint: You will need to use one of Excel’s Logical functions to ensure a 0 value for any 
suburb without an assigned shop code in this table. 
Distance is the distance from each assigned franchisee store to each suburb. The current 
assignment of suburbs to the franchisee is indicated in the template. 
Actual subscribers are different from prospective subscribers. A prospective subscriber is the 
likely maximum number of FHG subscribers in the suburb indicated by market research. The 
actual number of subscribers is dependent on the number of prospective subscribers and their 
distance from the nearest store. Prospective subscribers are identified in the Distance Survey 
and Suburb Profile Table of the Constants sheet. 
Actual subscribers are the number of prospective subscribers reduced by 10% for every 1 whole 
kilometre away (rounded down) from the nearest store until there are 0 actual subscribers. For 
example, a suburb with 62 prospective subscribers that is 4.8 kilometres away from the nearest 
store would have 38 actual subscribers. Mathematically, this can be represented as (𝑎𝑎𝑎𝑎𝑎𝑎𝑎𝑎𝑎𝑎𝑎𝑎
𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠𝑠 = 62 − ⌊62 × ⌊ ⌋ × 0.1⌋). 
Essentially – the further away prospective subscribers are from the store, the fewer actual
subscribers there will be.
Subscription Annual Revenue is the number of actual subscribers in each suburb according to 
the assigned store multiplied by the subscription rate for each subscription type (14, 21, or
42weeks subscriptions) multiplied by the number of weeks for subscription in a year. 
Hint: Although subscribers may take up the service, or may drop the service, or change 
subscription, you should assume that such changes cancel each other out – that is, you 
may assume that the number of subscribers does not change over the twelve-month 
period and all customers renew their subscription for the period (or, those that leave are 
replaced by new customers). Here is a hint that you can use the FLOOR and MAX 
functions.
Hint: To do this last requirement (row 55), you need to add column totals for each column 
in Current Franchise Distribution. 
You should then complete the Summary Table of Current Customers, Distance, and 
Revenue by store as indicated in the template and then sort the table by Annual Revenue
(largest to smallest), then by Distance (largest to smallest), then by Customer (largest to 
smallest). The Revenue per Kilometre is the Total Revenue for all stores divided by Total 
Distance for all stores.
# Franchise Redistribution Sheet 
As mentioned, Harvey is very environmentally conscious and wants to reduce greenhouse gas 
emissions and increases FHG’s environmental credentials. Rather than the previous random 
arrangements, Harvey wants to ensure that all suburbs are serviced by the closest FHG store. 
You are to use the Solver analysis in this worksheet to work out a possible redistribution plan 
to distribute groceries from the four franchisees to all suburbs, ensuring that the total distance 
of stores from the suburbs assigned to them is minimised. 
For ‘variable area’ in Solver analysis, use the area highlighted in light orange on this sheet 
in the template. These variables are binary (0 or 1) and are used to assign suburbs to stores. 
You will need to use the Simplex LP Solving Method. 
In this sheet, each row represents a suburb, and each column headed with a shop code (Columns 
B, C, D & E) represents the assigned shop (the highlighted light orange area). In the 
intersecting cell of the shop and the suburb, a 0 indicates that the shop is not assigned to the 
suburb, whereas a 1 indicates that the shop is assigned to the suburb. 
The key constraints are that each inner-city suburb should be assigned to one, and one only, 
store. Further, the variable area (the light orange cells) is either 0 or 1 (i.e., binary). The 
solver should be used to assign each suburb to its nearest store. The solver solution (i.e., 
original values) in the template is the current franchise distribution. 
Note: Please note that your formulas in this solution should be efficient. You can use a 
Lookup and Reference function to achieve this. 
After you fill in the total Franchisee assigned column, identify the current shop code
assigned to each suburb name and identify the newly assigned shop code for each suburb 
name. Using this code, identify the New Shop Location (i.e., the suburb of the assigned shop) 
in the next column. 
The Distance column shows the distance between the currently-selected shop to the assigned 
suburb (i.e. the cell in the light orange matrix with a ‘1’ in it). You should calculate this using 
an efficient formula. 
Hint: Consider both the ‘Distance to Shop Calculation Matrix’ and the Franchise 
Redistribution Solver solution. 
 
Hint: Remember that you used a Math & Trig function in Distance Survey and Suburb
Profile Table to determine the distance from each suburb to each shop. The Distance to 
Shop Calculation Matrix is automatically filled once your Distance Survey and Suburb 
Profile Table in the Constant Sheet is completed. 
Then, identify the subscriber numbers (14, 21 and 42), and total revenue based on this 
arrangement. Note that these actual subscriber numbers are calculated according to the
same formula outlined in the Current Franchise Distribution sheet (i.e., actual subscribers 
are calculated based on the number of prospective subscribers and the suburb’s distance from 
the new assigned store - the nearest store). 
Hint: You will need to use some of Excel’s Lookup & Reference, Math & Trig, and 
Logical functions in some of your formulas. 
Calculate the total figures for these columns at the bottom of the table. 
To easily identify the suburbs and franchisees that require changes, use Conditional 
Formatting to highlight (background only) the ‘Suburb Names’ in red (Column A) for each 
row if the assigned store stays the same, and highlight (background only) the ‘Suburb Name’ 
in green (Column A) if the assigned store changes. 
Similarly, use Conditional Formatting to highlight (background only) the ‘New Shop Code 
Assigned’ (Column H) cells for each row in Bright Orange if the assigned store changes. 
Save the results of Solver to a new answer sheet and restore the original values before 
submitting. 
Note: It is important that you restore the original values after running Solver. 
Note: Copy the original matrix values for the highlighted yellow section from the original 
sheet if you overwrite these values in error. 
Finally, you should complete the summary table of the revenue by store as indicated in the 
template. Identify the suburb name of each shop code using a formula, and the remainder of 
the summary table can be completed using Mathematical and Trig function. Do not forget to 
sort table by the same order as the previous sheet. 
The Revenue per Kilometre figure is the Total Revenue of all stores divided by Total Distance 
of all stores.
# Further Analysis using Pivot Table and Pivot Chart 
You are to create a Pivot Table and from that you create a professionally formatted Pivot Chart 
using the information on the Franchise Redistribution Sheet. You should put the Pivot Table 
and the Pivot Chart in two separate worksheets. 
Hint: These worksheets do not exist in the template.
When creating the Pivot Table, set the Suburb Name as a filter on the Pivot Table Fields so 
that the Pivot Chart can be modified to focus on the selected suburbs according to the viewer’s 
wishes. The row labels of the Pivot Table should be the four actual suburbs in which stores 
are located, and there should be three columns that calculate the number of subscribers for 
each of the three subscription types for each store. You should also include another column 
for total revenue. 
Hint: You do not need to use calculated fields in the PivotTable, which means that you 
do no need to add columns/rows manually in PivotTable. You should make these happen 
by customizing the ‘PivotTable Fields’ panel. 
Then you create a Pivot Chart (including the chart title, axes titles, colours, data labels, etc) 
as a Combo Chart Type on its own worksheet. 
Hint: For the Combo Chart type, consider Line Chart & Cluster Column Chart. You 
should also set up a secondary axis. 
The Chart should show the new shop location, the name of the suburb on Filter, the number of 
subscribers of each subscription types, and total revenue to allow them to be compared. 
Submit this Pivot Chart with all data shown (i.e. all shops with all suburbs assigned to them 
shown on the graph). 
Note: You should put the Pivot Table and the Pivot Chart in two separate worksheets 
and give each of the two worksheets a meaningful sheet name. 
Note: The pivot table should be edited through either the Pivot Table Fields or Pivot 
Chart Fields in Excel.
# New Store Investment Sheet 
FHG is a growing business. Harvey is thinking to build a new franchisee store at Rocklea in 
responds to his thoughts on expanding their presence at the Rocklea Grocery Markets in a few 
years’ time. This would allow him to acquire groceries directly and more cheaply, thus being 
able to provide more benefits to the customers. 
Harvey wants to save the same amount each term (monthly or fortnightly) for building a new 
store several years from now that will cost a certain amount of money. He investigates several 
scenarios of money deposit to undertake this investment plan. The monthly / fortnightly deposit 
would be used to cover the fix cost and variable cost of building a new store in a few years. 
You are to use the Excel Scenario Manager to create a Scenario Summary for each of the 
following scenarios: 
• Medium Case: Fix cost of $ 90,000, Variable cost of $16,000, Annual interest rate for 
saving of 1.6%, 12 monthly deposits each year over 5 years, Period start date is from 
15/02/2023. 
• Best Case: Fix cost of $ 95,000, Variable cost of $13,500, Annual interest rate for 
saving of 2.1% and fortnightly (26) deposits each year over 3 years, Period start date is 
from 15/02/2023. 
• Likely Case: Fix cost of $ 105,000, Variable cost of $12,000, Annual interest rate for 
saving of 2.7% and 12 monthly deposits each year over 6 years, Period start date is 
from 15/02/2023.
• Worst Case: Fix cost of $ 115,000, Variable cost of $14,000, Annual interest rate for 
saving of 3.5% and 12 monthly deposits each year over 7 years, Period start date is 
from 15/02/2023. 
Use the Scenario Manager to include a worksheet on your Workbook that contains the Scenario 
Summary. This is to summarise the four different scenarios. 
Note: The cells in Column B of the Output Area are the Result Cells for the Scenario 
Summary. 
You should add meaningful row labels to the scenario summary. 
Hint: This means that you need to copy the labels in the Input Area and the Output Area 
to the appropriate row in the Scenario Summary sheet. 
To assist with the calculation, you must complete the Deposit Schedule for all deposits 
identified in the scenarios. 
Note: This schedule extends from cell D5:I5 to as far as you need to go to accommodate 
the full deposit schedule for all scenarios considered. 
The Deposit # 1 has been entered for you. You should enter formulas starting from Cell D7
to display 2, 3, 4 … for Deposit # instead of using auto fill series. You should enter formulas 
start from E6 to display each period start date instead of manual input. 
As the scenarios vary in the number of deposits, the length of the schedule will need to be long 
enough to accommodate the highest number of deposits # in the scenarios. For example, the 
Medium Case scenario has 12 monthly deposits each year over 5 years = 60 rows, the Best 
Case scenario has 26 payments over 3 years = 78 rows, and the Worst Case scenario has 12 
monthly payments each year over 7 years = 84 rows. 
The opening balance for the first deposit should be zero. The closing balance is the amount 
collected after the interest earned has been added to the deposit for each term. The closing 
balance of one deposit period is the opening balance of the next deposit period. 
You should use opening balance for each deposit and term interest rate to calculate term deposit 
interest. Then you are required to use Excel’s Financial Function, PMT, to calculate deposit 
for each term. 
Hint: The amount of deposit for each term should be the same. You need to use Days 
function to get the period start date automatically.
 
Hint: Your ‘Deposit Schedule’ should display deposit amount as positive numbers. 
Hint: The last cell of your schedule should be exactly equal to the total amount of cost – 
that is, after your last deposit the closing balance should be equal to the total amount of 
cost of building the new store. 
You should professionally but simply format the Deposit Schedule. This means that all rows 
of the schedule that have values in them should have borders. 
Hint: Refer to Font – ‘All Borders’ format. 
This means that you should set your formulas so that the Deposit Schedule does not display 
rows when payments are finalised (i.e. no more rows after the closing balance equals to zero).
Note: Do not provide totals for the columns of the Deposit Schedule as this information 
will be displayed in the Output Area. 
Hint: You can force rows not to display the results of functions by using the ‘If’ function 
in your formulas to set the cell value to “” if the final deposit has been done. 
 
Hint: You are to use Conditional Formatting choosing one of the “Rule Type” so as to not 
show borders of cells that have no values in them. 
You will need to calculate the information in the Output Area using data entered in the Input 
Area and calculated in the Deposit Schedule. 
From the Input Area, you can calculate the Total Amount of Cost and Number of Deposits. 
From the Deposit Schedule you can calculate the Deposit Amount for Each Deposit, the 
Total Value of Deposit Made for building the new store, and Total Interest Collected Over Life 
of Deposits. These three items should be calculated as positive, not negative, numbers.
# Comments to Harvey Sheet 
In undertaking this extensive analytical exercise, there are three comments you wish to raise 
with Harvey. You must address the following three points: 
1. Considering the results of your analysis in the Current Franchise Distribution Sheet and 
the Franchise Redistribution Sheet, identify and discuss a single weakness of the Excel 
Model that relates to the business impact of the proposed redistribution plan. 
 
You may discuss any weaknesses that relates to the actual Excel model developed or to possible 
practical business problems that you identify with the proposed redistribution plan. You must 
identify the weakness and answer why you think it is a weakness. 
Note: Do not write a paragraph – 3 to 5 lines would be enough to identify the weakness 
and explanation on why you think it is a weakness. 
Note: The weakness should be related to the Excel business model rather than Excel tool 
itself or Excel functions. 
2. In the Current Franchise Distribution and the Franchise Redistribution sheets, you 
calculated a ‘Revenue Per Kilometre’ figure. This is a business metrics that can reflect FHG’s 
performance. Please identify another 3 business metrics that indicates a business’s growth or 
decline. 
3. Harvey wants to know more about the Excel solution you have built for his business. 
Please explain to Harvey in detail that how Predictive Business Analytics is utilised in this 
Excel model? 
Hint: Predictive Business Analytics is the use of statistics and modelling techniques to 
determine future performance based on current and historical data. It can be used to 
improve Business Performance with Forward-Looking Measures. 
 
Note: Do not write a paragraph – 3 to 5 lines would be enough to explain how Predictive 
Business Analytics is utilised in this Excel model.
## 作业代写价格:
|类型|美元|人民币|
|-----:|-----:|-----:|
|Assignment|$40-$120|¥400-¥800|

### 为了方便快速定价和确定是否代做，麻烦提供私聊的时候提供以下信息：
如果是作业，提供本次作业要求文档；如果是考试，提供考试范围和考试时间
一两句话概括一下作业任务或者考试任务，比如”可以帮忙实现一个决策树算法吗？”、”网络安全选择题考试，范围是内网横向移动知识点”
### 作业定价有两种方式：
    - 根据定价标准进行
    - 通过微信我们一起协商
## 联系方式
[为您提供 CS , DS , 商科 编程作业代写](http://dzuoye.work "编程代写")
微信（WeChat）：design2023866

<img src="design2023866.jpg"  width="200" />

