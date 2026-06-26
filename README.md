# International Brand Product Catalog Scraper

> Python automation scripts for collecting structured product information from official international brand websites and preparing catalog data for e-commerce import.

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue)](https://www.python.org/)
[![Data Output](https://img.shields.io/badge/Output-Excel%20%2F%20WooCommerce-success)]()
[![Use Case](https://img.shields.io/badge/Use%20Case-Product%20Catalog-orange)]()
[![Status](https://img.shields.io/badge/Status-Active-brightgreen)]()

---

## Overview

This repository contains Python-based product catalog scraping tools designed to collect product information from official international brand websites.

The scripts help transform product URLs into structured catalog data that can be reviewed, edited, and imported into an e-commerce platform.

The workflow focuses on product information such as:

* Product name
* Brand
* Product model
* Product description
* Technical specifications
* Product category
* Product images
* Product documents or datasheets
* SEO metadata
* WooCommerce-ready Excel output

The source code is available in the [`script`](../../tree/script) branch.

---

## Supported Brand Scripts

| Script          | Brand  | Main Purpose                                                          |
| --------------- | ------ | --------------------------------------------------------------------- |
| `lutron_uji.py` | Lutron | Collect product details and generate WooCommerce-ready catalog output |
| `atago_uji.py`  | ATAGO  | Collect product information from official ATAGO product pages         |

> Each scraper may use different extraction rules because every brand website has a different page structure.

---

## Key Features

### Product Information Extraction

The scripts can collect structured product information from official product pages, including:

* Product title
* Brand name
* Product model
* Description
* Full product content
* Technical specifications
* Category or breadcrumb information
* Main product image
* Datasheet or document links when available

### Excel-Based Workflow

The workflow starts from an Excel file containing product URLs.

The script processes each URL and produces an output Excel file that can be reviewed before product import.

### WooCommerce Catalog Output

The generated Excel output can include fields such as:

```text
Name
Brand
Product Description
Product Short Description
Product Categories
Product Tags
Meta Description
Image URL
Source URL
Product Specifications
Processing Status
```

### Image URL Detection

Some international brand websites hide product images inside dynamic page elements, JavaScript components, or CSS backgrounds.

The scripts use additional extraction logic to locate the most relevant product image URL when it is not directly visible in the HTML.

### Content Cleaning

The scripts may clean unwanted website elements before generating the output, such as:

* Navigation text
* Footer content
* Cookie banners
* Social media labels
* Repeated page elements
* Irrelevant language fragments
* Generic landing page content

### Optional Browser Automation

Some websites load product information dynamically.

For these cases, the scripts can use browser automation to access content that is not available through standard HTTP requests.

---

## Workflow

```text
Excel Product URL List
        ↓
Read Product Page
        ↓
Extract Product Information
        ↓
Clean and Validate Content
        ↓
Detect Product Image
        ↓
Generate Structured Catalog Data
        ↓
Export Excel File
        ↓
Review Before Import
```

---

## Installation

Clone this repository:

```bash
git clone https://github.com/Pradana123/Script-Katalog.git
cd Script-Katalog
```

Switch to the script branch:

```bash
git checkout script
```

Install the required Python packages:

```bash
pip install pandas openpyxl requests beautifulsoup4
```

For websites that need browser rendering:

```bash
pip install playwright
playwright install
```

---

## How to Use

### 1. Prepare an Excel File

Create an Excel file containing product URLs.

Example:

| Name                | Brand  | URL                                    |
| ------------------- | ------ | -------------------------------------- |
| Digital Light Meter | Lutron | https://example-brand.com/product-page |
| Refractometer       | ATAGO  | https://example-brand.com/product-page |

The minimum required field is usually:

```text
URL
```

Additional fields such as product name, brand, and model can be used as fallback data when the product page does not provide complete information.

---

### 2. Run the Script

Example for a Lutron catalog scraper:

```bash
python lutron_uji.py --input "lutron_products.xlsx" --output "lutron_catalog_output.xlsx"
```

Example with request delay:

```bash
python lutron_uji.py --input "lutron_products.xlsx" --output "lutron_catalog_output.xlsx" --delay 0.6
```

Check all available options:

```bash
python lutron_uji.py --help
```

For the ATAGO script:

```bash
python atago_uji.py --help
```

---

## Output

The script generates an Excel file containing structured product catalog information.

Typical output includes:

```text
Product Name
Brand
Model
Product Description
Short Description
Specifications
Product Category
Product Tags
SEO Metadata
Image URL
Source URL
Processing Status
Website Validation Notes
```

The output should be reviewed before importing into WooCommerce, WordPress, or another e-commerce platform.

---

## Intended Use

This repository is designed for:

* Product catalog preparation
* E-commerce product migration
* WooCommerce product import
* Technical product data collection
* Internal product database development
* Product information standardization
* SEO content preparation
* Marketplace catalog management

---

## Responsible Use

These scripts are intended for collecting publicly accessible product information from official brand websites.

Before using any script, please ensure that you:

* Respect the website's Terms of Service
* Review the website's robots.txt policy
* Use reasonable request delays
* Avoid excessive traffic
* Do not bypass login systems or access restrictions
* Do not collect private, personal, or restricted data
* Verify all collected data before publishing it

This repository is not affiliated with, endorsed by, or officially connected to any brand referenced in the scripts.

All product names, trademarks, logos, images, and technical content belong to their respective owners.

---

## Project Structure

```text
Script-Katalog/
│
├── README.md
│
└── script/
    ├── lutron_uji.py
    ├── atago_uji.py
    └── additional scraping utilities
```

---

## Notes

* Website layouts can change at any time.
* A scraper may need updates when a brand changes its website structure.
* Product data should always be checked manually before publishing.
* Image URLs and technical specifications may vary depending on the source page.
* Some pages may require browser-based rendering before content can be extracted correctly.

---

## Author

Created and maintained by [Pradana123](https://github.com/Pradana123)

---

## Contact

For collaboration, improvement ideas, or technical discussion, please use the repository issue section or contact the repository owner through GitHub.
