================================================
  TikTok Live + Shop Report Builder
  Quick-Start Guide
================================================

WHAT THIS TOOL DOES
-------------------
This tool takes your TikTok data exports, combines them,
and produces a single Excel report with three sheets:

  - Product Comparison  (Live vs Shop performance per product)
  - Daily Overview      (Day-by-day shop + live data)
  - Live Summary        (All live sessions with totals)


STEP 1 — Export your files from TikTok
---------------------------------------

FILE 1: live_product.csv
  Where to get it:
    TikTok Live backend → Analytics → Product Performance
  Export as CSV. Rename the file exactly: live_product.csv

FILE 2: shop_product.csv
  Where to get it:
    TikTok Seller Center → Data → Product Analysis → Export
  Export as CSV. Rename the file exactly: shop_product.csv

FILE 3: live_performance.csv
  Where to get it:
    TikTok Live backend → Analytics → Live Performance
  Export as CSV. Rename the file exactly: live_performance.csv

FILE 4: shop_daily.csv
  Where to get it:
    TikTok Seller Center → Data → Overview → Daily Breakdown → Export
  Export as CSV. Rename the file exactly: shop_daily.csv

  NOTE: File names must be spelled exactly as shown above,
  including lowercase letters.


STEP 2 — Drop the files into the input folder
----------------------------------------------
Open this folder: nextt_pipeline/input_data/

Drag and drop all four CSV files into that folder.

You don't need all four files every time. The tool will
skip any sheet it can't build and tell you which file
is missing.


STEP 3 — Run the report
------------------------
Double-click:  run_report.command

A black Terminal window will open and run automatically.
When it says "Report saved to output/" it is finished,
and the output folder will open in Finder.


STEP 4 — Find your report
--------------------------
Open this folder: nextt_pipeline/output/

Your report is named:  live_report_YYYY-MM-DD.xlsx
(where YYYY-MM-DD is today's date)

Open it in Excel or Numbers.


TROUBLESHOOTING
---------------
"Missing file: [name]..."
  → The tool couldn't find that CSV in input_data/.
    Check the file name matches exactly (no extra spaces,
    correct .csv extension).

"Could not find a Python installation with pandas"
  → Python is not set up correctly on this computer.
    Ask your tech team to run:  pip3 install pandas openpyxl
    Then try again.

"Nothing happens when I double-click run_report.command"
  → Right-click the file → Open → click Open in the
    dialog that appears. You only need to do this once.

The report opened but columns look wrong / data is missing
  → Make sure you exported the correct report from TikTok
    (see Step 1) and that the file name is exactly right.


QUESTIONS?
----------
Contact your data team or the person who set up this tool.

================================================
