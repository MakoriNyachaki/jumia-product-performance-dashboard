# Building a Jumia Product Performance Dashboard in Excel.

## Introduction

Jumia and its sellers need to understand how price, promotions, and customer feedback influence product performance. In this project, the number of reviews is used as an engagement proxy because the dataset does not include units sold or revenue. Do not describe reviews as sales or claim that a relationship proves causation.

We will analyze Jumia's product performance to translate raw data into meaningful insights that drive business growth and improve service delivery. By examining customer engagement, pricing models, and promotional effectiveness, we can better understand the specific factors influencing product success.

## Business Questions

1. Are larger discounts correlated with a higher volume of reviews?
2. Do highly-rated products naturally attract stronger customer engagement?
3. What is the relationship between product pricing and customer ratings?
4. Which products represent the top performers based on ratings and review volume?
5. Which products require a pivot in their current pricing or marketing strategy?

## Tools

- Excel - For cleaning, analyzing, and visuliztion and reporting
- Power Query - Load and clean the dataset

## Data Cleaning

*Note:* Copy your dataset and retain the original. Work on the copied dataset.

We found the following errors and inconsistencies in the dataset:

- Mispelt headings _Ratingd_ instead of _Rating_
- Mismatching Header cases _old prices_ and _Current prices_
- Trailing white spaces and extra spaces in column 1
- Duplicates: we found 3  duplicates with an exact match in all fields.
- Negative review values in the reviews column
- One price range in both the _Current Price_ and _Old Prices_ columns instead of a price value.
- Improper data types in the _Rating_ and _Prices_ columns.
- Missing values in the `Rating` and `Review` columns.

So, after copying our dataset into the new sheet to retain the original dataset, create a data table. Select all the data, go to *Insert > Table* from the _Tables group section._

![Image for the raw data](/images/raw-data.png)
*Fig 1: A photo showing original uncleaned data*

### Cleaning Decisions

The following are the steps and decisions that we made through the cleaning process.

1. For the headings, we had to change spellings into correct spellings for _Rating_ from _Ratingd_
2. For the price columns, we had to convert headers using `PROPER()` function to convert them to proper case. That is, the first letter of each word is capitalized.
	`
	=PROPER(B1)
	=PROPER(C1)
	`
_B1_ and _C1_: is the address to the _Current Price_ and _Old Price_ headers in the sheet. The function returns the new text formatted to proper case. 
Then we used copy and paste special(Past as Values) where we copied and pasted the new text to the header replacing the old one.

3. Removed all trailing white space and  extra spaces using the `TRIM()` function. 

*Steps*

* Insert a new column to the right of the _Product_ column
* In Cell B2, type the function below:

	`=TRIM(A2)`

* Press _Enter_ to return the new text value
* The column is auto filled.

_Copy and Paste special_

* Select all the values in the B column i.e. from B2:B116
* Copy the values using the Shortcut `CTRL + C`
* Then, select cell A2, and right- click on it
* Then from the dialog box choose, _Paste as Values_ instead of _Paste_.

Use the `CLEAN()` function to also remove non-printable characters from the _Product_ column.

*Steps*

* Clear the column we used for trimming
* In _B2_ enter the following function
	`=CLEAN(A2)`
* Copy the values in _B2:B116_, paste special in in cell A2.

4. Duplicates

Three duplicates were found; they were an exact match in every field. Remove them to avoid redundancy and incorrect results of analysis due to redundancy.

*Steps*
* Select all the data
* On Data Tab click on _Remove Duplicates_ in the _Data Tools_ group section
* Selcect all the fields, and then click _OK_

A confirmation pop-up message will show you the number of duplicates removed.

5. Neagtive Reviews

For the negative values in the _Reviews_ column, convert them to absolute values using the `ABS()` function.

*Steps*

* Insert a new column to the right of the _review_ column
* On the new column, F, enter the following function in cell _F2_:
	`=ABS(E2)`

E2 is the address of the first review negative value
* Press Enter
* Copy and Paste special the values to replace those in the _review_ column.

6. Price Ranges

Used the midpoint of the price range to get the value for the particular products. We found the midpoint using the formula:

	`Price = (Lower Range Limit + Upper Range Limit) / 2`

7. Data Types

*Current Price and Old Price Columns*

* Used Find and replace to replace _KSh_ with Blanks
* Converted the remaining values to _Number_ type, then to _Currency_ type to 2 decimal places with the symbol _KSh_, and used the *,* 1000 separator 

*Rating*

Converted Rating to _Number_ type formatted to one decimal place.

*Review*

Converted _Review_ column to _Number_ type, a whole number.

*Steps to Change to Number Type*

* Select the Column with the data
* Go to Home tab, on the Number group section, from the Drop Down, select _More Number formats_
* On the left side of the pop-up window, select _Number_ or _Currency_ and specify the other options.
* Then click _OK_ to apply the changes.

8. Missing Values

For the missing values and blanks, never assume anything. Leave them blank as they are. Note that a _Blank_ does not mean zero.

Exclude blanks from metrics that require the missing field.

Add a data status field (for example, _Complete_, _Missing Rating_) if it can help users understand better.

## Formulaas and Enriching Fields

We added the following columns to our cleaned data.

*1. Discount Amount*

Use the following formula to get the values for the column:

`=@[Old Price]-@[Current Price]`

Optionally, compare the advertised discount with the calculated discount:

`=IFERROR(([@[Old Price]]-[@[Current Price]])/[@[Old Price]],"")`

Flag any difference instead of overwriting the advertised value.

*2. Rating Category*

Rating values are classified as follows:

*Poor* < 3

*Average* Between 3 and 4

*Excellent* > 4.5

`=IF([@Discount]="","Missing",IF([@Discount]<20%,"Low Discount",IF([@Discount]<=40%, "Medium Discount","High Discount")))`

4.1-4.5 are not classified; if the original boundaries must be followed literally, label 4.1-4.5 as `Unclassified`.

*3. Discount Category*

Classify your _Discount_ into the following categories:

_Low Discount_ < 20%

_Medium Discount_ 20%-40%

_High Discount_ > 40%

`=IF([@Discount]="","Missing",IF([@Discount]<20%,"Low Discount",IF([@Discount]<=40%,"Medium Discount","High Discount")))`

*4. Price Category*

- On the `Analysis` worksheet, define thresholds rather than arbitrary values. A preferred approach uses the first and third quartiles

	`
		=QUARTILE.INC(tblProducts[Current Price],1)
		=QUARTILE.INC(tblProducts[Current Price],3)
	`

- Name those sales as `Price_Q1` and `Price_Q3,` respectively.
- Use the following classifications

`=IF([@[Current Price]]="", "Missing",IF([@[Current Price]]<=Price_Q1,"Low Price",IF([@[Current Price]]<=Price_Q3, "Medium Price","High Price")))`

- Remember to record the final KSH thresholds in the `Data Dictionary` for ease of interpretation by other users.

![Image for the cleaned data](/images/cleaned-data.png)
*Fig 2: A photo showing cleaned data*

## Analysis and Pivot Tables

![Image for the Pivot Tables and Charts](/images/pivot-tables.png)
*Fig 4: A photo PivotTables and PivotCharts*

## Dashboard Features

```text
+------------------------------------------------------------------+
| Title, refresh date, and slicers                                 |
+------------------------------------------------------------------+
| Total Products | Avg Price | Avg Discount | Avg Rating | Reviews |
+------------------------------------------------------------------+
| Top 10 by Rating | Top 10 by Reviews | Top 10 by Discount        |
+------------------------------------------------------------------+
| Discount vs Reviews | Rating vs Reviews | Price vs Rating         |
+------------------------------------------------------------------+
| Rating Mix | Discount Mix | Key Insights / Recommendations       |
+------------------------------------------------------------------+
```

![Image of the complete dashboard](/images/dashboard.png)
*Fig 3: A photo showing the Jumia Product Performance Dashboard*

## Findings

- High discounts do not mean a high number of reviews or better engagement. From the analysis, there is no relationship at all between discount and reviews. The correlation coefficient between discount and reviews is *-0.14*.
- There is a weak relationship that establishes no pattern between rating and reviews. This indicates that more reviews do not translate to higher ratings in the data; hence, no causation. The correlation coefficient stands at 0.06.
- The analysis establishes a weak positive correlation between prices and ratings of 0.11. Prices are not influenced by the rating of the product.

## Recommendations

- Investigate quality issues on highly discounted products but have low ratings.
- Find out why highly rated products have higher prices than the averagely and lowly rated products. Does this influence come from sales made? 
- Combine promotions of products with high ratings and higher prices to increase engagement and check if they can translate to sales.

## Limitations and Lessons

Learnt:

- Using and applying formulas and functions
- Created a dashboard
- How to clean data, build pivot tables, and pivot charts
- How to report and communicate insights and findings.

Limited to:

- Time constraints, since more research and a deeper understanding is critical to a deeper dive and understanding.

## File Structure and opening the Workbook

```text
jumia-product-performance-dashboard/
├── .gitignore
├── README.md
├── data/
│   └── Excel_jumia_dataset.csv
├── dashboard/
│   └── jumia_product_dashboard.xlsx
└── images/
    ├── raw-data.png
    ├── cleaned-data.png
    ├── pivot-tables.png
    └── dashboard.png
```
*To open the workbook:*
1. In the project folder navigate to the `dashboad/` subfolder
2. Double click on the `jumia_product_dashboard.xlsx` to open using *Excel*.
