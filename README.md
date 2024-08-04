# xll_emma

This add-in grabs data off the [EMMA](https://emma.msrb.org/) website
and pulls it into Excel.
Download [`xll_emma.xll`](?) and open it in Excel.
It only works on Excel for Windows and you may get some warnings
about enabling macros. Madge, you are soaking in the
source code and can build it yourself if you are nervous.

The list of available curves is returned by `EMMMA_ENUM()`.
Provide one of those as a first argument to `EMMA` and
and optional date argument. The default is the most recent data
available.

A two row array of time in years and rates is returned.