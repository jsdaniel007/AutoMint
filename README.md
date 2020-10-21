# AutoMint
With Mint as the data source, Excel as the data storage, and openpyxl for data manipulation and logic; organize, store and visualize personal financial transaction data.

#### Planned Features:
- History: Long-term storage record of mint excel export, only adding new transactions into a record organized by months and years
- Lifetime Data/Visualization: Calculates spending data by category, Needs vs. Wants, Excess funds, and potential visualizations like pie chart
- Snapshots: Keep a 3-month snapshot with it's own calculations similar to Lifetime Data stats/visualizations
- Grocery Shopping Optimizer: compile grocery info such as average cost per grocery run, have a self evolving list of optimized grocery list based on ... (TBD)

#### Project Pipeline:
Mint --> Excel Export --> openpyxl Processing --> Excel Data Manipulation
