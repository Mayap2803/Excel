# Excel
Personal project using a excel pivot tables and VBA macros to analyse the most read Fiction books of 2024
<div class='tableauPlaceholder' id='viz1734180497112' style='position: relative'><noscript><a href='#'><img alt='Pie Chart ' src='https:&#47;&#47;public.tableau.com&#47;static&#47;images&#47;DD&#47;DDCNQG7HT&#47;1_rss.png' style='border: none' /></a></noscript><object class='tableauViz'  style='display:none;'><param name='host_url' value='https%3A%2F%2Fpublic.tableau.com%2F' /> <param name='embed_code_version' value='3' /> <param name='path' value='shared&#47;DDCNQG7HT' /> <param name='toolbar' value='yes' /><param name='static_image' value='https:&#47;&#47;public.tableau.com&#47;static&#47;images&#47;DD&#47;DDCNQG7HT&#47;1.png' /> <param name='animate_transition' value='yes' /><param name='display_static_image' value='yes' /><param name='display_spinner' value='yes' /><param name='display_overlay' value='yes' /><param name='display_count' value='yes' /><param name='language' value='en-GB' /><param name='filter' value='publish=yes' /></object></div>                <script type='text/javascript'>                    var divElement = document.getElementById('viz1734180497112');                    var vizElement = divElement.getElementsByTagName('object')[0];                    vizElement.style.width='100%';vizElement.style.height=(divElement.offsetWidth*0.75)+'px';                    var scriptElement = document.createElement('script');                    scriptElement.src = 'https://public.tableau.com/javascripts/api/viz_v1.js';                    vizElement.parentNode.insertBefore(scriptElement, vizElement);                </script>
### Aims:
To create a database and investigate trends in the most popular fiction books read in 2024. 

### Process:
Information for the Top 20 Most Read Fiction Books by Week in 2024 information was taken from [Amazon Charts](https://www.amazon.com/charts/2024-01-07/mostread/fiction?ref=chrt_bk_nav_fwd)

#### Creating the Database:
An excel table was created to store the following information: Starting Week date, Ranking in the Week, Book Title, Author, Number of Reviews, Average Review Rating, Number of Weeks in The Top 20, If Avalibale on Kindle Unlimited, Genres. The starting week began on 31st December 2023 and ended on the week begining 24th November 2024. 

#### Analysis:
Used a combination of SUM() and COUNTIF() formulae to return the total number of unique book titles throughout the time period and found 58 unique book titles with the most read books being written by Rebecca Yarros, J.K. Rowling, and Sarah J. Maas, all of which were on the Top 20 list every week.

![No Unique_Books](https://github.com/user-attachments/assets/4f23b1ec-5c1f-4735-9cd6-ca3bde456b1e)

The entire database table was transformed into a pivot table in order to be able to filter the data and create charts to visualse the number of weeks each book was on the [Top 20 bar chart](https://github.com/user-attachments/files/18026989/Total_weeks_on_top_20.pdf). A bar chart showing the number of times a book written by each author made the weekly Top 20 list is shown below. 

![top10 authors](https://github.com/user-attachments/assets/da294dcd-ea08-4e2a-bd8b-155aeecfb3d0)





