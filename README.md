### Overview of Project
The Stock Analysis project is all about helping my friend Steve who has asked me to analyze a green stock for his parents to see if it is worth investing in. Steve's parents are very much passionate about green energy. So they decided to invest in DAQO New Energy Corp. DAQO's ticker symbol is "DQ". So further we'll be using the ticker symbol "DQ" in our analysis.

- As mentioned we'll first start with analysing the DQ stock :

## DQ Analysis
- As Steve's parents are starting to pester him about DAQO's stock we'll be starting our analysis with "DQ" .
- So here we are using VBA code to further help Steve in stock analysis. 
- I first started with creating Macros and write my code under that particular macro which was DQ Analysis and below is the screenshot for the same :

![image](https://user-images.githubusercontent.com/92283185/140091905-ec5085d4-d0d3-46dc-8569-9ad4fa423695.png)

- From the above code i tried to analyze the DQ stock for Steve and his parents and came to conclusion that the DQ doesnt have good returns in 2018 which we'll further see in the results section below.

## All Stock Analysis
- Since Daqo might not be the best option for Steve's parents to invest in, I started analying multiple stocks to find some better choices for them.
- With a little more code, we can analyze a whole list of stocks as follows: 
![image](https://user-images.githubusercontent.com/92283185/140094806-511e3395-0204-4dde-9773-11e289600bf5.png)
![image](https://user-images.githubusercontent.com/92283185/140094930-9d754094-36e0-4081-9d86-bfd3fce77791.png)

We'll analyze its results in the further section of my analysis.
 
## Results
As we discussed in earlier section that in order to  help Steve and his parents to calculate the total daily volume and yearly return we started working first on stock "DQ" (for which Steve's parents were interested in) and further did all stock analysis. This resulted into as follows :

![image](https://user-images.githubusercontent.com/92283185/139620519-45622d77-7993-4ee8-814f-aba0cc4d5b4e.png)


![image](https://user-images.githubusercontent.com/92283185/139620614-00cf97d0-695e-464f-b768-8b35b172c62a.png)
## Summary
- To summarize the Stock Analysis project i refactored the original code. The following Questions and explanation for the same will help Steve and his parents to understand what stock to invest in and what should be the correct way of doing it. We'll see some examples below with coding and screenshots.
- In order to make my code more efficient, I needed to switch the nesting order of my for loops. To do this, I created a 4 different arrays; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays with the tickers array by using a variable called the tickerIndex.
So my Refactored code looks like this:

![image](https://user-images.githubusercontent.com/92283185/140096875-c53dda67-2c02-4074-a3ff-dbe961688f1e.png)
![image](https://user-images.githubusercontent.com/92283185/140096991-0714870f-cf77-4a15-a1e2-0afcd910ffad.png)
![image](https://user-images.githubusercontent.com/92283185/140097154-83205d0a-38e9-4383-873a-eedb07f7175b.png)

And then if i compare my 2018(original) run time and 2018(refactored) run time my refactored code runs about 20 seconds faster than the original code making it more efficient.
![image](https://user-images.githubusercontent.com/92283185/140109947-7d2d3081-379c-4c5f-baaf-77969cc0d48e.png)
![image](https://user-images.githubusercontent.com/92283185/140110641-733f0eb7-d046-454f-9208-83f8ea8bfedd.png)


# What are the advantages or disadvantages of refactoring code?
- The advantage of refactoring code is the screen running time is less then the original script and the refactoring is more easier to understand and read.
- But the only disadvantage i relaized is it is time consuming and very complex to **"write"** the script 
- As we saw earlier the **refactored code** is more tedious. 
- This is further understood by following screenshot where we can cleary see that time to run 2017(refactored one) is less as compare to 2018(original one).
- <img width="631" alt="VBA_challenge_2018" src="https://user-images.githubusercontent.com/92283185/140099668-74f35023-b831-4a91-99e9-d2fc555ca8ec.png">
<img width="598" alt="VBA_challenge_2017" src="https://user-images.githubusercontent.com/92283185/140099725-9954cb59-bcb9-4d72-9e57-60a0c1158316.png">

# How do these pros and cons apply to refactoring the original VBA script?
If i talk about the pros and cons of stock analysis from original to refactoring the most important pros I noticed was 
- The resources were utlizied less as compare to the original vba script.
- And at the same time the cons was the refactoring is more complex structure.
- for e.g In Original VBA script in order to calculate the total daily volume and its returns 
      - It was more time consuming and resources were over utilized while 
      - When we calculated the same total daily volume and returns for 2017 it took comparatively less time which we noticed in the screenshot earlier
      - To run the code it took 0.97 as compared to 2018 where it took 1.07 seconds) which is self explanatory.








