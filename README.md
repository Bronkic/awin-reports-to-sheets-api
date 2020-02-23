This tool is meant to facilitate reportings and analysis for Advertiser and Agencies in the Affiliate Marketing business.

Most reporting tools out there offer a variety of data sources that let you  automatically integrate data from services like Google Analytics, Facebook, Instagram and many more into your reports.
However, only very few offer a data source for AWIN, the biggest Affiliate Network in Europe. And all of them require you to pay for them. 

What most reporting tools, like Google Data Studio or Swydo, do however offer is Google Sheets as a data source. This means if you download statistics from your AWIN Advertiser account, import them into a Google Sheets Spreadsheet, and then use a framework to arrange the data as needed, you don't need to manually input numbers into your reporting tool anymore.

However, this still requires you to manually download and import specific statistics every time you need to create a reporting. This is why I created this script. It can automate the downloading and importing (using cron jobs), effectively allowing you to completely automate the creation of all your AWIN reportings. For an agency responsible for monthly reportings to a number of Advertisers, this can save a lot of time. Additionally, you could even schedule this tool to be executed daily or even hourly. This way you could create your own live BI dashboard.

This tool also checks whether statistics that you have imported previously still contain open sales/leads and updates them, if they have been validated by now!

Do note that this was my very first project, so be gentle. Although I have tested the tool thoroughly, I cannot guarantee an absence of bugs. It's also very possible that there are much smarter and more efficient ways to implement certain features.

Also, AWIN has a limit of 20 requests per minute. So running this tool could sometimes take a while, if you have a high number of programs and/or open sales to update.

Setup:
1. Get OAuth2-Authentication-Token from AWIN and enter it into the yellow cell of the first sheet (Token) of setup.xlsx: https://wiki.awin.com/index.php/Advertiser_Service_API
2. Get Authentication for Google Sheets and Google Drive APIs and add token.pickle and credentials.json to the same folder as the tool: https://developers.google.com/sheets/api/guides/authorizing
3. Create a Google Sheets Spreadsheet and copy its Spreadsheet-ID (the long string of characters at the end of the URL)
3. Add your Awin Advertiser Accounts (Awin Advertiser ID, Name and Spreadsheet-ID) to the second sheet in Setup.xlsx.
4. Run the program.
