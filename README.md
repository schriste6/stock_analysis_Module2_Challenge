# Stock Analysis - Refactored Code, a gain or a loss?

## Overview of Project:  

### Original Request
*Customer, Steve, is pleased with the original **Stock Analysis** code developed to analyze stock performance, it delivers on all the requested deliverables.  The initial request was a stock analysis code to run on 12 stocks.  Steve's delight with the original deliverable has him returning with a change request.*
### Change Request
*Steve would like to run the **Stock Analysis** on thousands of stocks. The concerned is the current code will not deliver needed results in a timely manner.  The new request is to refactor the current code to deliver the stock analysis on thousands of stocks with improved runtime.*  

### Results: 
- Snippets of code runtimes for 2017 (refactored on left / original on right) 
  - Refactored code demonstrates an ***improved runtime of 0.55859 seconds***:

![](/Resources/VBA_Challenge_2017.png)     ![](/Resources/Original_code_2017_code_performance.png)

- Snippets of code runtimes for 2018 (refactored on left / original on right) 
  - Refactored code demonstrates an ***improved runtime of 0.95898 seconds***:

![](/Resources/VBA_Challenge_2018.png)     ![](/Resources/Original_code_2018_code_performance.png)

#### Refactored code can be accessed via this [link](/Resources/Refactored_script_w_improved_runtime.txt)
- Adding a tickerIndex simplified the code:
  - eliminated a for loop 
  - reduced number of conditionals
  - reduced memory needs 

#### Summary:  
##### What are the advantages or disadvantages of refactoring code?
- Advantages of refactoring code include but are not limited to; improved organization, improved perferformance, improved quality.
- Disadvantages of refactoring code include but are not limited to; risk of introducing bugs, risk the developer may misunderstanding the purpose, risk of corrupting complex/big scripts.  
##### How do these pros and cons apply to refactoring the original VBA script?
- In this example, the script is basic and specific to one customer.  The advantages are a big gain.  
