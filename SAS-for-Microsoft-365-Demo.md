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

6. 4.	Open Excel and create a new workbook. Select SAS Viya in the ribbon along the top of the window. 
Note: In this demonstration we use the Microsoft desktop applications, however SAS for Microsoft 365 works with the web applications as well.

5.	In the SAS Viya tab, select Home. A panel opens on the right, prompting you to login to SAS Viya. In the SAS-provided virtual lab, enter Student as the User ID and Metadata0 as the password. Once you are connected, the panel provides access to all reports, data, and programs in the SAS Viya environment.

6.	In the Home tab, select the Favorites drop-down list. You can quickly access open, recent, and favorite SAS reports and data. Click   (More options) to view additional actions, such as the ability to upload data to SAS Viya and customize preferences.


<img width="1922" height="1032" alt="Image" src="https://github.com/user-attachments/assets/34159138-15d2-4c92-8679-def1c816f28d" />

------------------------------------------------------------------------

# Excel: Insert Report Content

9.  In the SAS pane, select **Reports**.
10. Expand **My Favorites** and double-click **Warranty Analysis**.
11. Drag the left side of the pane to make it larger.
12. Verify the **Cost Overview** page is active.
13. Click **2016** to filter the data.
14. Maximize the bar chart.
15. Select **Insert in Document**.
16. Resize the image and click **Update**.

------------------------------------------------------------------------

# Excel: Insert Data

17. Select the **Data** tab.
18. Navigate to:

All tables \> SAS Studio compute context server \> SASHELP \> HOMEEQUITY

19. Configure:

-   Insert rows → All rows
-   Remove columns: YOJ, DEROG, DELINQ, CLAGE, NINQ, CLNO
-   Filter: BAD = 1
-   Sort by LOAN

20. Insert into new worksheet named **HOMEEQUITY Data**.

------------------------------------------------------------------------

# Excel: Insert Program Results

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

You can insert report objects into: - Outlook emails - PowerPoint
slides - Word documents

Use **Insert in Document** or **Attach report as PDF** from the SAS
pane.

------------------------------------------------------------------------

For full documentation, click **Help** in the SAS Viya ribbon.
