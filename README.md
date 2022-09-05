# Analysis of Results

## Overview of the Project

### Purpose

The purpose of this project is to analyze stocks and find the better option to invest in. While the original code allows us to look at a different set of stocks and their performance for specific years, it may not be efficient when it comes to larger data. When analyzing, we need to make sure the code can run for hundreds of stocks. To do that, I refactored the original analysis.


### The code

The purpose of this analysis is to find the stock performance by finding total daily volumes and the returns from each stock. Thus, the importance of this code lies in accessing the right data from our spreadsheet rows.

The more significant part of the code includes finding the volumes for each stock:


<img width="635" alt="Volume" src="https://user-images.githubusercontent.com/111609994/188519244-1eb0eedf-a220-4164-8787-6b8539762ac6.png">

In addition, in order to calculate the return, my code needed to access the starting and ending price for the stock:

<img width="620" alt="Screen Shot 2022-09-05 at 4 00 17 PM" src="https://user-images.githubusercontent.com/111609994/188519326-66299194-8937-4455-9e8c-d59b470e4b98.png">

## Results

### Original vs. Refactored

As mentioned above, the purpose of refactoring the code is to improve the code and its execution time. Here we'll see that the original code works slower compared to the refactored code. In addition, the refactored code will allow the user to loop over many more stocks than the original code would be able to. 

*Here are some visuals comapring the execution times between original and refactored code from 2017 and 2018:

**Original Code Run for 2017**

<img width="372" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/111609994/188519698-14b6ab14-ab24-4aa3-a387-88895761b038.png">

**Refactored Code Run for 2017**

<img width="372" alt="VBA_Challenge_2017 Refactored" src="https://user-images.githubusercontent.com/111609994/188519707-610a02c4-4bb5-4a14-bb7d-c61772a7860e.png">

**Original Code Run for 2018**

<img width="372" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/111609994/188519724-57ab9cca-1ab6-4a32-93a2-e07f4d9593ac.png">

**Refactored Code Run for 2018**


<img width="372" alt="VBA_Challenge Refactored_2018" src="https://user-images.githubusercontent.com/111609994/188519754-1d3f8e3f-5a2d-42bd-8ea6-b46e498d6bf6.png">

### 2017 vs. 2018

After running the code it becomes obvious that throughout 2017 the returns were much higher than 2018. Stocks performed way better. 

**Stock Performance From 2017**

<img width="347" alt="Return for 2017" src="https://user-images.githubusercontent.com/111609994/188519506-c7337c76-f9ac-477d-9f41-6faa85ce17d6.png">

**Stock Perfomance from 2018**

<img width="355" alt="Return for 2018" src="https://user-images.githubusercontent.com/111609994/188519535-afc7f42c-1218-4382-9d17-8b5d0584dc44.png">


# Refactoring the Code

## Advantages and Disadvantages

Refactoring a code is similar to editing the code in order to make the code more understandable. The advantages of refactoring include making it more comprehensible for others. In addition, the code becomes shorter, includes cleaner methods, and eventually restructuring the code in the future won't impact its functionality. Refactoring allows us to get rid of duplications and fix redundancies.

Some disadvantages of refactoring include the introduction of many more errors and bugs if one isn't careful. Refactored code is not self-evident thus it's encouraging to use comments to explain the text.


## Impact on original VBA script

Refactoring the original VBA script resulted in shorter text, thus less processing time for the code. We saw above that refactoring return our results faster. This was due to the simpler code. 
