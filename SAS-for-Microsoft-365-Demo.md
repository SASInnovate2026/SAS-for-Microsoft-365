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

6.  Open the **Microsoft Excel** desktop application and create a new workbook.

7.  Select the **SAS Viya** tab in the top ribbon.

8.  Select **Home** and log in with the following credentials:

    **User ID:** Student

    **Password:** Metadata0

    When asked if you want to opt in to all your assumable groups, select **Yes**.

![Login Page](<../OneDrive - SAS/Desktop/Innovate 2026/login1.png>)

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
