import pipeline


#A test list of Australian and American companies to monitor
tickers = [
    "FLT.AX", "JHX.AX",
    "AAPL", "MSFT", "GOOGL", "AMZN", "TSLA", "META", "NFLX", "NVDA", "JPM", "V", "MA", "AMD",
    "INTC", "CRM", "ADBE", "PYPL", "DIS", "UBER", "SBUX", "NKE", "COST", "PEP", "KO"
]

test = pipeline.ASXPipeline(tickers, "exampleWorkbook.xlsx", initialiseWorkbook=False)

"""Perform some test plots"""

#Plot 1 year Opens for Apple vs Uber 
test.plot(["AAPL", "UBER"]) 

#plot Amazon, Apple and Costco volatility
test.plot_volatility(["AMZN", "AAPL", "COST"])

# Return the correlation between Google and Microsoft prices (We expect a strong correlation)
print(test.correlate("MSFT", "GOOGL"))

#Return a summary of UBER performance
print(test.statSummary("UBER"))

#Compare daily returns of Apple against ASX200
test.plotReturnsAgainstASX200("AAPL")