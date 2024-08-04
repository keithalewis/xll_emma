# xll_emma

This add-in grabs Municipal Yield Curves and Indices off the 
[EMMA](https://emma.msrb.org/ToolsAndResources/MarketIndicators) 
website and pulls them into Excel.
Download [`xll_emma.xll`](?) and open it in Excel.
It only works on Excel for Windows and you may get some warnings
about enabling macros. Madge, you are soaking in the
source code and can build it yourself if you are worried.

The list of available curves is returned by `EMMMA_ENUM()`.
Provide one of those as a first argument to `EMMA` and
and an optional date argument. The default is the most recent data
available. It returns a two row array of time in years and rates.
Use the Excel function 
[`TRANSPOSE`](https://support.microsoft.com/en-us/office/transpose-rotate-data-from-rows-to-columns-or-vice-versa-3419f2e3-beab-4318-aae5-d0f862209744)
to get a two column array.

Only Bloomberg provides both non-callable and callable curves.