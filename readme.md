# crypto_vba
An Excel/VBA project to communicate with various cryptocurrency exchanges APIs. Tested on Windows 10 & Excel 2016, but should work for Excel 2007+.

# Exchanges:
Get information from/send information to:
- [Binance](http://binance.com/)
- [Bitfinex](https://www.bitfinex.com/)
- [Bitstamp](https://www.bitstamp.net/)
- [Bittrex](https://www.bittrex.com/) 
- [Coinbase](https://www.coinbase.com)
- [CoinbasePro](https://pro.coinbase.com/)
- [Coinone](https://coinone.co.kr/)
- [Coinspot](https://www.coinspot.com.au/)
- [HitBTC](https://hitbtc.com/)
- [Kraken](https://www.kraken.com/)
- [Kucoin](https://www.kucoin.com/)
- [OKEx](https://www.okex.com/)
- [Poloniex](https://www.poloniex.com/) 
- [Coinigy](https://www.coinigy.com/) - not an exchange, but a service where you can access multiple exchanges for a fee - not actively maintained
- [Cryptopia](https://www.cryptopia.co.nz/) - WARNING: exchange suspended following a hack
- [GDAX] -> see CoinbasePro
- [Liqui] -> exchange closed
- [WEXnz] -> exchange closed, removed

The API response is pure JSON, for which I included https://github.com/VBA-tools/VBA-JSON to process and a function to build on that.
As excel/VBA development is not very compatible with GIT, my pushes/forks/updates might be clunky.
Please consider the code I provide as simple building blocks: if you want to build a project based on this code, you will have to know (some) VBA. There are plenty of courses available online, two simple ones I send starters to are: https://www.excel-pratique.com/en/ and https://homeandlearn.org/.

# How to use?
Import the .bas files you need or simply take the sample Excel file. In the modules you'll find some examples how to use the code. Feel free to create an issue if things don't work for you. The project uses quite some Dictionaries in VBA, check out e.g. https://excelmacromastery.com/vba-dictionary/ if you want to know a bit more about them.

# ToDo
- Excel formulas need better caching to prevent a stalling/crashing Excel
- Better error handling

# Done
- For historical prices, included https://www.cryptocompare.com/api/ (now https://min-api.cryptocompare.com/ )
- Build excel functions to get the information directly to a sheet, has some caching, but - TEST PHASE
- Working examples of several exchanges in the example file
- Created a basic XLSM sample file for all provided exchanges
- ArrayToTable improvement to handle various data types (e.g. Trade and Margin trade) in one JSON response
- Post-process the Array to a more usable format (flat table)
- Process the response to something you can use in Excel: an array/Range etc.
- Build a function to transform the JSON to an Array
- Build tests for all modules/functions
- Integrate VBA-JSON into the project
- Build the Binance API connector
- Build the Bitfinex API connector
- Build the Bitstamp API connector
- Build the Bittrex API connector
- Build the Coinbase API connector
- Build the GDAX/CoinbasePro API connector
- Build the Coinone API connector
- Build the Coinspot API connector
- Build the Cryptopia API connector -> not updated yet after the hack of Jan 2019
- Build the HitBTC API connector
- Build the Kraken API connector
- Build the Kucoin API connector
- Build the OKEx API connector
- Build the Poloniex API connector
- Build a working and tested VBA hash function
- Build a function to transform Dictionaries into JSON and URLencode
- Added the UrlEncode function for e.g. Cryptopia (and Excel versions before 2016)
- Removed inactive exchanges: Liqui, WEXnz/BTCe (nostalgia, that was the first exchange i got working in excel)

# Donate
If this project/the Excel saves you a lot of programming time, consider sending me a coffee or a beer:<br/>
BTC: 1DNFF9y3dDMLNURpgdT3wXmFpmGBsQRyPa <br/>
ETH (or ERC-20 tokens): 0x9070C5D93ADb58B8cc0b281051710CB67a40C72B<br/>
DOGE: DHSN2ZEaLqoSW6v9Mg39pwNktHBD7ESSsi <br/>
<b>Cheers!</b>