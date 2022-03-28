## Introduction
Amazon provided its fulfillment centers with a variety of internal web-based tools, developed and maintained by professional software developers, that leaders could use to pull close-to-real-time data on site productivity, quality, and safety metrics.  However, operational leaders at individual warehouses often needed these data tailored or processed in ways beyond what these tools are able to offer.  A common format that warehouse-level analysts like myself would use to accomplish these extra processing steps in a semi-automated, portable way was the macro-enabled Excel dashboard, enabling leaders to click a button and have the most recent data populate in a digestible format.

Broadly speaking, the VBA macros I built fall into three categories:

### Paste and Run
paste data and then run the code (split list, pick order formulas, vna consolidation formulas). only interacts with cells
(under construction)

### Web Scraping
scrape and transform most recent realtime data (capacity, consolidation, pallet audit, iol). interacts with cells and webpages
(under construction)

### Web Scraping + Shared Drives
scrape most recent data, combined with shared drive file from an ETL job - not as portable (pallet pull, damageland). interacts with cells, web, and files
(under construction)
