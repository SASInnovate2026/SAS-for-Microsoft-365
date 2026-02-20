# SAS for Microsoft 365 Demo Guide

Bringing together SAS Viya and Microsoft 365 allows you to take
advantage of the strengths of both worlds. You can use SAS Viya to
prepare your data and create deep analytical insights, then explore and
share results in familiar Microsoft applications including Excel, Word,
PowerPoint, and Outlook.

------------------------------------------------------------------------

# Table of Contents

# Notes for me

-   Add a slide explaining who has access to SAS for Microsoft 365.
-   Provide installation/loading instructions.
-   Consider including a QR code linking to documentation or workshop
    resources.
-   Add SVG icons and screenshots where helpful.
-   Upload workshop instructions to GitHub for easy access.

------------------------------------------------------------------------

# Part 1: Open the Warranty Analysis Report in SAS Visual Analytics (edit once I have the image)

1.  Open **SAS Visual Analytics** from the Applications menu by
    selecting **Explore and Visualize**.

2.  Navigate to:

    SAS Content \> Products \> SAS Visual Analytics \> Samples

3.  Double-click **Warranty Analysis** to open the report.

4.  Select **Opened Reports** → **Close all reports**.

5.  Right-click **Warranty Analysis** and select **Add to Favorites**.

------------------------------------------------------------------------

# Excel: Access the SAS Panel

6. Open The Microsoft Excel desktop application and create a new workbook. Select **SAS Viya** in the ribbon along the top of the window. 
**Note:** In this demonstration we use the Microsoft desktop applications, however SAS for Microsoft 365 works with the web applications as well.

5.	From the SAS Viya tab, select **Home**. A panel opens on the right, prompting you to login to SAS Viya. In the SAS-provided virtual lab, enter **Student** as the User ID and **Metadata0** as the password. Once you are connected, the panel provides access to all reports, data, and programs in the SAS Viya environment.

<img width="1922" height="1032" alt="Image" src="https://github.com/user-attachments/assets/34159138-15d2-4c92-8679-def1c816f28d" />

6.	On the **Home** tab, select the **Favorites** drop-down list. You can quickly access open, recent, and favorite SAS reports and data. Click **More options** to view additional actions, such as the ability to upload data to SAS Viya and customize preferences.

**screenshot here**

------------------------------------------------------------------------

# Excel: Insert Report Content

7.	From the SAS pane, select the **Reports** tab. Expand **My Favorites** and double-click the Warranty Analysis report.
 **screenshot**

9.	Drag the left border of the SAS pane to enlarge the report preview. In the **Results** tab, you can view the different report pages and interact with the report objects, just like in Visual Analytics.  For example, verify the Cost Overview report page is active. Click **2016** to filter the data and update the report objects.

**screenshot**
10.	Hover over the bar chart and click **Maximize**. Alternatively, you can right-click the chart and select Maximize view.
**screenshot**
11.	Select either **Object menu**or right-click the chart and select **Insert in document**. An image of the bar chart is added to the Excel worksheet, along with an indicator of the filter selection.
**screenshot**

12.	Select the graph image in the worksheet. From the SAS Viya tab in the ribbon, click **Selected Object** to update the object, remove it from the document, or unlink it from SAS.
**screenshot**

14.	Resize the graph to make it wider, then click **Selected Object > Update** or **Update All** from the SAS Viya tab to update the image based on the new dimensions.
**screenshot**

16.	In the report preview in the SAS pane, select **2015** in the button bar. Right-click the graph and select **Update in document** to refresh the graph based on a different subset of the data.
**screenshot**

18.	The detailed table below the bar chart can also be inserted into Excel. First, select a cell in column A of the worksheet below the graph. This will determine where the next object will be inserted. Then right-click the table in the SAS pane and select **Insert in document**.
**screenshot**

20.	The table can be further enhanced in the worksheet by using Excel features. Highlight the header row in the report, then select Home in the Excel ribbon. Click **Sort &** **Filter > Filter**. Use the option menu on Gross Labor Amount to select **Sort Largest to Smallest**. Next, select the data values under Gross Labor Amount and select **Conditional Formatting > Data Bars** and click the orange gradient fill option.
**Note:** Depending on the Excel modifications made in the worksheet, the changes may or may not persist after updating report objects.  
**screenshot**

22.	Hover over the bar chart in the SAS pane and click **Restore** to return to the full report view.

23.	We have added two report objects, but you can also add all objects from a page. Select the **Labor Details** page to view the three visualizations. Then right-click anywhere on the page and select **Insert active page in document**. A new worksheet is created named Labor Details, which includes the three charts.


------------------------------------------------------------------------

# Excel: Insert Data
18.	Next, select the Data tab in the SAS pane. You can view tables defined in your SAS Viya environment and insert data into Excel for local exploration and processing. Select All tables > SAS Studio compute context server, then select the SASHELP library and the HOMEEQUITY table.

19.	A new Results tab is now active, providing an opportunity to subset the data and select table options prior to inserting it into the Excel worksheet.

1.	In the General tab, change Insert rows to All rows.  

2.	In the Columns tab, clear the check boxes for the following columns: YOJ, DEROG, DELINQ, CLAGE, NINQ, and CLNO.

3.	In the Filter tab, you can either build an expression graphically, or code the filter expression from scratch. Click Graphical builder and confirm the selected variable is BAD. Click   (Filter) then type 1 in the Value field. When BAD=1, it indicates the loan defaulted. Note that you can also use the Lookup Value button to select values from the input data. Click OK.  

4.	In the Sort tab, click Add sort and change the column to LOAN.

20.	Create a new worksheet in the Excel file and rename it HOMEEQUITY Data. Click the first cell in the upper left corner, then in the SAS pane click   (Insert table in document). The selected data is added to the spreadsheet. You can take advantage of Excel functionality to enhance the view with filtering, sorting, and formatting.  

21.	Not only can you view SAS data in Excel, but there is also an option in the SAS Viya tab in the ribbon to Upload Data. That provides the opportunity to prepare data in Excel, then load it to SAS Viya for further analysis in applications like SAS Studio, Visual Analytics, or Model Studio. **do this if time**


------------------------------------------------------------------------

# Excel: Insert Program Results

22.	Examine the HOMEEQUITY data more closely. Notice the column City is in lower case. This can be resolved temporarily in Excel, or you can use SAS code to create a new table with the proper case values. Select Programs and enter the following program in the Code tab. Notice as you type, autocomplete prompts provide suggestions for valid keywords and tables.
data work.homeequity_update;
   set sashelp.homeequity;
   City=propcase(City);
run;
This DATA step will capitalize the first letter of each word in City.

23.	Additional SAS syntax may be used to summarize or analyze data. Add the following PROC FREQ step to generate frequency counts for the HOMEEQUITY_UPDATE table.
proc freq data=work.homeequity_update  order=freq;
 tables State / nocum;
run;
24.	Click Run. The Results tab includes the frequency report. Create a new worksheet and select the upper left cell. Click   (Insert in document) to add the results.

25.	What if the program is modified? Return to the Code tab and add a TITLE statement to provide a custom title and ODS NOPROCTITLE statement to hide the procedure name.
26.	title "Number of Loans by State"; 
27.	
28.	ods noproctitle;
29.	proc freq data=work.homeequity_update order=freq;
...
30.	Click Run. In the SAS panel, click   (Update in document) to refresh the results in the worksheet.

31.	Click Output Data to view the new table created by the program. Scroll to the right to confirm the City column is in proper case. Click   (Open) to access the Results tab. Here you have the option to filter and sort the data and insert the results into the Excel workbook.

32.	Return to the Programs tab. There is a note indicating that the program is embedded in the document. This means that if the Excel workbook is saved, closed, and reopened, it will contain active links to your SAS content, as well as any embedded SAS programs. To save the program outside of the Excel file so that it is accessible in other applications or to your colleagues, click   (More options) and select Save. You can name the file and choose a storage location in SAS Content. Click Cancel.

33.	Click   (More options) and notice there is also the ability to open a program, allowing you to access and execute code from an external file.


``` sas
data work.homeequity_update;
   set sashelp.homeequity;
   City = propcase(City);
run;

proc freq data=work.homeequity_update order=freq;
   tables State / nocum;
run;
```

Insert results into a new worksheet.

------------------------------------------------------------------------

# Outlook & PowerPoint

NOTE: Outlook is not configured in the SAS Virtual Lab to complete this portion of the demonstration. It is recommended that you watch the video.
30.	Open Outlook to see some of the unique features available when sharing SAS content via email. Create a new email message, then click SAS Home to access the SAS pane.

31.	In the Home tab, change the drop-down list to Favorites and double-click the Warranty Analysis report. You may also select Reports to access any reports in SAS Content.

32.	In the Cost Overview report page, maximize the Bar Chart object. Click   (Object menu) for both the bar chart and the table to select Insert in message. Note that the menu also allows you to attach individual report objects as a PDF.

33.	Select   (More options) in the upper-right corner of the Results tab, then select Attach report as a PDF. There are many choices available to customize the document.

34.	In the Options tab, verify the following options are selected:

o	Show page numbers
o	Include Table of Contents
o	Include cover page

35.	In the Select Objects tab, clear the check boxes for Detailed Cost Analysis and Labor Group Analysis and Forecast. Click OK. Open the attached PDF to view the result.  

36.	Finally, you can add a link to the live report in SAS Visual Analytics. In the email message, type Link to full report:. Then click   (Insert report link in message). Additional text or information can be added before sending the email message.

PowerPoint and Word: Insert Report Content
You can also insert report objects into Microsoft Word documents and PowerPoint slides. The features in both applications are similar, so we will demonstrate in PowerPoint.
37.	Open PowerPoint and create a new blank slide deck. Right-click the title slide and select Layout > Title Only. Type Cost Overview in the Title box.

38.	From the SAS Viya tab in the ribbon, select Home and log in. In the Home tab in the SAS pane, select Favorites > Warranty Analysis. Expand the SAS pane to enlarge the view of the report.

39.	In the Cost Overview page, find the Bar Chart object and select   (Object menu), then click Insert in Document. Notice that the bar labels are staggered in the graph because there isn't enough space to display them on the same line. Resize the graph to fill the slide, then click   (Object menu) and select Update in document to regenerate the bar chart in the allocated space.

40.	Create a new slide and change the layout to Blank. Insert the Cost Change from Previous Year by Primary Labor Group bubble plot. Again, resize the image to fill the slide, then click   (Update all objects in document). You may use all PowerPoint functionality to create beautiful, insightful presentations, with embedded, live results from your SAS reports. 
We've toured several of the applications that are supported with SAS for Microsoft 365. To learn more, click Help in the SAS Viya ribbon to access complete documentation.  

