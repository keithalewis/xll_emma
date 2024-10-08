# xll_emma

This add-in grabs Municipal Yield Curves and Indices off the 
[EMMA](https://emma.msrb.org/ToolsAndResources/MarketIndicators) 
website and pulls them into Excel.
Download [`emma.zip`](emma.zip), extract all files, then open `emma.xlsx` and `emma.xll` in Excel. 
You may have to right click on `emma.xll`, select Properties, and click Unblock if you see warnings.

The list of available curves is returned by `EMMMA_ENUM()`.
Provide one of those as a first argument to `EMMA(id, date)` and
and an optional date argument. The default is the most recent data
available. It returns a two column array of time in years and par coupons.

`EMMA.URL(id)` returns the URL of the raw curve data.  
`EMMA.HELP(id)` returns the URL of the EMMA help page for the curve.  
`EMMA.DATE(id, date)` returns the date of the most recent curve prior or equal to date.

Queries are cached to a SQLite database `emma.db` in the same directory as the add-in.

Only Bloomberg provides both non-callable and callable curves.
