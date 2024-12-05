# Excel
Personal project using a excel pivot tables and VBA macros to analyse the most read Fiction books of 2024

### Aims:
To create a database and investigate trends in the most popular fiction books read in 2024. 

### Process:
Information for the Top 20 Most Read Fiction Books by Week in 2024 information was taken from [Amazon Charts](https://www.amazon.com/charts/2024-01-07/mostread/fiction?ref=chrt_bk_nav_fwd)

#### Creating the Database:
An excel table was created to store the following information: Starting Week date, Ranking in the Week, Book Title, Author, Number of Reviews, Average Review Rating, Number of Weeks in The Top 20, If Avalibale on Kindle Unlimited, Genres. The starting week began on 31st December 2023 and ended on the week begining 24th November 2024. 

#### Analysis:
Used a combination of SUM() and COUNTIF() formulae to return the total number of unique book titles throughout the time period and found 58 unique book titles with the most read books being written by Rebecca Yarros, J.K. Rowling, and Sarah J. Maas, all of which were on the Top 20 chart every week.
![No Unique_Books](https://github.com/user-attachments/assets/4f23b1ec-5c1f-4735-9cd6-ca3bde456b1e)

The entire database table was transformed into a pivot table in order to be able to filter the data and create charts to visualse the number of weeks each book was on the Top 20 chart.
[Total_weeks_on_top_20.pdf](https://github.com/user-attachments/files/18026989/Total_weeks_on_top_20.pdf)





