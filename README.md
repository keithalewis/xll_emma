# xll_emma

This add-in grabs Municipal Yield Curves and Indices off the 
[EMMA](https://emma.msrb.org/ToolsAndResources/MarketIndicators) 
website and pulls them into Excel.
Download [`emma.zip`](emma.zip), extract all files, open `emma.xlsx` and `emma.xll` in Excel. 
You may have to right click on `emma.xll`, select Properties, and click Unblock.

The list of available curves is returned by `EMMMA_ENUM()`.
Provide one of those as a first argument to `EMMA` and
and an optional date argument. The default is the most recent data
available. It returns a two column array of time in years and par coupons.

Queries are cached to a SQLite database names `emma.db` in the same directory as the add-in.

Only Bloomberg provides both non-callable and callable curves.