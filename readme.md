# crypto_vba
An Excel/VBA project to communicate with various cryptocurrency exchanges APIs

# First steps:
- [kraken](https://www.kraken.com/)
- [poloniex](https://www.poloniex.com/) , first steps
- [btc-e](https://www.btc-e.com/) , first steps

The API response is pure JSON, for which you can use e.g. https://github.com/VBA-tools/VBA-JSON to process, which might be included here in this project.
As this is my first Git experiment and Excel/VBA and git don't work that well together, my pushes/forks/updates might be clunky...

# How to use?
Import the .bas files you need: starting with ModHash.bas and adding the exchange(s) you might need. In the modules you'll find some examples how to use the code. Feel free to create an issue if things don't work for you.

# ToDo
- Build the Poloniex and BTC-e APIs
- Better error handling
- For historical prices, include https://www.cryptocompare.com/api/
- Process the response to something you can use in Excel: an array/Range etc.
- Build an XLSM file with working examples