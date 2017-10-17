# crypto_vba
An Excel/VBA project to communicate with various cryptocurrency exchanges APIs. Tested on Windows 10 & Excel 2013/2016, but should work for Excel 2007+.

# Exchanges:
Get information from:
- [kraken](https://www.kraken.com/)
- [poloniex](https://www.poloniex.com/) 
- [Bittrex](https://www.bittrex.com/) 
- [Liqui] (https://liqui.io/)
- [btc-e](https://www.btc-e.com/) *OUTDATED due to end of BTC-e


The API response is pure JSON, for which I included https://github.com/VBA-tools/VBA-JSON to process and a function to build on that.
As this is my first Git experiment and Excel/VBA and git don't work that well together, my pushes/forks/updates might be clunky...

# How to use?
Import the .bas files you need or simply take the sample Excel file. In the modules you'll find some examples how to use the code. Feel free to create an issue if things don't work for you.

# ToDo
- Expand the XLSM file with working examples
- Better testing
- Better error handling
- For historical prices, include https://www.cryptocompare.com/api/
- Build excel functions to get the information directly to a sheet
- Later: place orders

# Done
- Created a basic XLSM sample file0
- ArrayToTable improvement to handle various data types (e.g. Trade and Margin trade) in one JSON response
- Post-process the Array to a more usable format (flat table)
- Process the response to something you can use in Excel: an array/Range etc.
- Build a function to transform the JSON to an Array
- Build tests for all modules/functions
- Integrate VBA-JSON into the project
- Build the Bittrex API connector
- Build the BTC-e API connector
- Build the Poloniex API connector
- Build the Kraken API connector
- Build the Liqui API connector
- Build a working and tested VBA hash function

# Donate
If this project/the Excel saves you a lot of programming time, consider sending me a coffee or a beer:<br/>
BTC: 1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa <br/>
ETH: 0x6f61c0d77f410e614e294c454380bbb6ecc7bdc1<br/>
<b>Cheers!</b>