# Excel VBA Order Processing Automation

> **Project status:** Work in progress

## Project Overview

This portfolio project demonstrates an Excel VBA solution for processing promotional customer orders. It automates data import, eligibility checks, duplicate detection and the preparation of fulfillment files.

The project is based on a realistic business scenario in which a weekly CSV export contains the complete historical order dataset rather than only newly submitted orders.

## Business Problem

Manual processing of a large order export is time-consuming and increases the risk of inconsistent decisions, duplicate fulfillment and data-quality errors.

The solution is designed to:

- Identify newly submitted orders
- Apply validation and eligibility rules
- Detect historical and current-batch duplicates
- Separate eligible and non-eligible orders
- Generate structured fulfillment outputs
- Reduce repetitive manual work

## Main Business Rules

- Only orders with a German delivery address can be eligible.
- The invoice-validation result is provided by the source system.
- A previous valid and fulfilled order blocks a later duplicate claim.
- Within the current batch, the first valid submission remains eligible and later duplicate submissions are rejected.
- An earlier invalid submission does not block a later valid submission.
- Duplicate detection uses normalized customer and address information.
- Customer name alone is not sufficient to identify a duplicate.

## Planned Workflow

1. Import the complete CSV export.
2. Identify records that have not been processed previously.
3. Normalize names and address information.
4. Apply validation, country and duplicate-detection rules.
5. Assign an internal fulfillment decision.
6. Prepare product-specific fulfillment files.
7. Export the final results.

## Tools and Skills Demonstrated

- Microsoft Excel
- VBA automation
- CSV data processing
- Business requirements analysis
- Business-rule implementation
- Data normalization
- Duplicate detection
- Test-case design
- Data-quality and process-risk analysis

## Repository Contents

The completed repository will include:

- Macro-enabled Excel demonstration workbook
- Synthetic CSV test dataset
- VBA source-code modules
- Screenshots
- Business-rule documentation
- Test scenarios and expected results

## Data Privacy

This repository contains only fictional and synthetic demonstration data. It does not contain real customer, invoice, address or company data.

## Key Business Analysis Insight

Customer-and-address matching may create false positives if promotional entitlement is intended to apply per purchase or per device rather than per customer or household.

The business must therefore define whether eligibility is based on the customer, household, invoice, purchase or individual device. A transaction- or device-level identifier would provide more deterministic duplicate detection.
