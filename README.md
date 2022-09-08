# google-trends-using-pytrends

Python script to collect google trends data via API with pytrends, automating data collection and analysis.

##Overview

Here, in the excel file there's list of keywords to search on google trends from 3 months before of the given date to 3 months after the given date.

Using pytrends we can access google trends API to get the data automatically and manipulate the recieved dataset with pandas dataframe. Then save the data to excel file.

##Note
Google API, specially google trends API blocks API calls too frequently for security reasons. So set time delay in between API calls or use PROXY IPs verified by Google.
Otherwise, your IP will be blocked for several hours on suspicion of being mailicious bot.
