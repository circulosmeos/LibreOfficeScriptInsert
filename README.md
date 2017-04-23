**LibreOfficeScriptInsert.py** inserts a python script (Macro) into a LibreOffice Calc document.   
   
Tested with LibreOffice 5.2, at least.
    
Use
===

     $ python LibreOfficeScriptInsert.py myDocument.ods myScript.py

A copy of *myDocument.ods* will created, with the script *myScript.py* embedded in it, and with name *myDocument.with_script.ods*

Example
=======

For example, the script [FIFOStockSellProfitCalculator.py](https://github.com/circulosmeos/FIFOStockSellProfitCalculator) that calculates FIFO profits, can be inserted into any of your .ods Calc documents for portability and ease of use, with:

     $ python LibreOfficeScriptInsert.py myStocks.ods FIFOStockSellProfitCalculator.py

The Calc document with the Macro script embedded will be named *myStocks.with_script.ods*

License
=======

Licensed as [GPL v3](http://www.gnu.org/licenses/gpl-3.0.en.html) or higher.   
