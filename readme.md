# crypto_vba
An Excel/VBA project to communicate with various cryptocurrency exchanges APIs. Tested on Windows 10 & Excel 2013, but should work for Excel 2007+.

# Exchanges:
Get information from:
- [kraken](https://www.kraken.com/)
- [poloniex](https://www.poloniex.com/) 
- [btc-e](https://www.btc-e.com/) 

The API response is pure JSON, for which I included https://github.com/VBA-tools/VBA-JSON to process and a function to build on that.
As this is my first Git experiment and Excel/VBA and git don't work that well together, my pushes/forks/updates might be clunky...

# How to use?
Import the .bas files you need: starting with ModHash.bas and adding the exchange(s) you might need. In the modules you'll find some examples how to use the code. Feel free to create an issue if things don't work for you.

# ToDo
- Build an XLSM file with working examples
- Better testing
- Better error handling
- For historical prices, include https://www.cryptocompare.com/api/
- Build excel functions to get the information directly to a sheet
- Later: place orders

# Done
- Post-process the Array to a more usable format (flat table)
- Process the response to something you can use in Excel: an array/Range etc.
- Build a function to transform the JSON to an Array
- Build tests for all modules/functions
- Integrate VBA-JSON into the project
- Build the Poloniex and BTC-e API connectors
- Build the Kraken API connector
- Build a working and tested VBA hash function