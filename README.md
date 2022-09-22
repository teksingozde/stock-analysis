# Stock-Analysis

## Overview of Project

All stock entry analyzes compiled depending on the year he entered were given to the customer at any time. These years are 2017 and 2018. This worksheet is included with the VBA script macros. The signs may change depending on the desired years. In addition, the total volume corresponding to the desired year and the corresponding values for that year are included in the worksheet. Each page contains 12 stocks.

### Purpose

A well-organized data set, ensuring the accuracy of the results, fast execution of macros and time savings are the criteria that every customer may want. An analysis should both provide clear interpretations and be easy to understand.

The main purpose of this study is to make macros more efficient and to save time when more rows of data sets or more stocks are added. A practical printout is provided with a single click and year entry. If such a VBA macro is not created, it will take a lot of time and decrease the accuracy in data analysis. The aim here is to be able to achieve great work in a short time.

### Results

A separately edited file was created from the original script and then tested for correct operation. For this study, the outputs of 2017 and 2018 are shown and comparisons are made.

<img width="282" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/26927158/191863951-bc39a23d-345f-498f-907b-ab3ae54edbeb.png">

<img width="277" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/26927158/191863961-f1411f89-d319-42fe-9c66-4cd6448ffa60.png">

###### Results for Refactored Stock Analysis

<img width="272" alt="VBA_Challenge_Refactored_2017" src="https://user-images.githubusercontent.com/26927158/191863969-dc2c223d-1b0a-45de-b761-72012d84a625.png">

<img width="275" alt="VBA_Challenge_Refactored_2018" src="https://user-images.githubusercontent.com/26927158/191863987-fa0508e2-6384-45f7-9580-31c02ba60ba8.png">

If we analyze the stock performance from 2017 to 2018, almost all of the 12 companies performed well in 2017. If we talk only for TERP, it gives the impression of a -7.2% drop. We can say that the best performances of that year are DQ with 199.4% and SEDG with 184.5%. Of course, we should not forget the ENPH and FLSR, which provide returns of over 100%.

Speaking for the 2018 analysis, there are two stocks that performed well. These are RUN with 84% and ENPH with 81.9%. In addition, the worst performing company was DQ with -62.6%. In fact, in order to make a correct analysis, diversifying the number of shares and data will give the most accurate results.

### Code and Comments

For a series of 12 stocks, a two-cycle with initial volume, starting and ending price, is created. The total volume is increased as much as the loss, within the matching of the current entry with the indexed stocks. There are also start and end values depending on the group order. In the VBA macro code, which you will see below, the stock, the total volume in the year and the annual return are written.

<img width="848" alt="Code_Screenshot1" src="https://user-images.githubusercontent.com/26927158/191864874-949e1aa4-85aa-47d4-8f4e-dad204023eb4.png">
<img width="887" alt="Code_Screenshot2" src="https://user-images.githubusercontent.com/26927158/191864875-1f526c60-fd4f-4278-b48b-c71827642b11.png">
<img width="891" alt="Code_Screenshot3" src="https://user-images.githubusercontent.com/26927158/191864876-80e9b9f4-14e9-486b-bfc2-b7de1605bb85.png">
<img width="888" alt="Code_Screenshot4" src="https://user-images.githubusercontent.com/26927158/191864877-2630e9a5-6f30-4b96-8002-f54d7418bc64.png">

## Summary 

##### 1.	What are the advantages or disadvantages of refactoring code?

Rearranging a program is often necessary for code fragments that function poorly. The situation we call improvement is done in order to save time in general terms, to be able to read better, and to increase the reliability rate of the analysis. In a successful edit, the code will run more accurately and the analysis will be done completely and with high reliability. Reorganizing the code means making improvements to the pieces of code in line with the wishes of the customers. In the research sector, data cleaning processes are carried out in line with the wishes and interests of the customers, we can call this a kind of optimization.

##### 2.	How do those pros and cons apply to refactoring the original VBA script?

While reorganizing VBA macros, it is decided to restructure the macros in order to increase the performance speed. In the reorganized piece of code, an accurate analysis will be possible in a short time when more datasets are added for this project or this piece of code is adapted to a larger volume dataset.




