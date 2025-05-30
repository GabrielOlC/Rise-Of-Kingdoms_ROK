### ROK App - Rise of Kingdoms KvK Data Analysis & Scoring Engine

> This repository contains the documentation and insights into the **ROK App**. It's objectiv is to manage and analyze player performance data for the game Rise of Kingdoms (ROK) during Kingdom vs. Kingdom (KvK) events.



### **Core Objective:**

To provide a robust, automated, and fair system for processing raw "kingdom scan" data, calculating comprehensive player scores across numerous metrics, computing contribution points (CP), and generating insights to aid in kingdom governance, player evaluation, and reward distribution.

**Key Features & Technical Implementation:**

* **Automated ETL Pipeline (Power Query):**
  
  * Ingests data from multiple Excel source files with specific naming conventions representing different scan periods.
  
  * Performs data cleaning, transformation, merging, and structuring to create a reliable dataset for analysis.
  
  * Handles common issues and inconsistencies found in raw game scan.

* **Scoring & Calculation Engine (Excel Formulas):**
  
  * Calculates player performance for T1-T5 troops (kills, deaths), Kill Points (KP), resource assistance, and other in-game metrics.
  
  * Implements a configurable "Achievements Contribution" metric, weighting player actions against customizable kingdom goals and player power levels.
  
  * Computes "Contribution Points" (CP) used for in-game ranking or reward purchases, with logic to handle points from farm/alt accounts and deductions for undesirable activities (e.g., zeroing/feeding).

* **Multi-Sheet Application Architecture:**
  
  * **Data Input:** Dedicated sheets for manual data entry not found in scans (e.g., Hall of Heroes data, pre-KvK stats, manual CP adjustments).
  
  * **Data Processing & Validation:** Intermediate sheets (Check Data Structure) display step-by-step calculations for each player and file block, facilitating error identification and data validation.
  
  * **Core Calculation Hub (Summed Data sheet):** A central sheet where all processed and manual data converges for final score and CP computation using an extensive array of interconnected formulas.
  
  * **Configuration (Settings sheet):** Allows kingdom leadership to customize goals, scoring metrics, weighting factors, and other parameters for each KvK event.

* **Error Detection & Reporting:**
  
  * Multiple layers of error checks, from Power Query transformations to specific Excel formula logic designed to flag inconsistencies or illogical data
  
  * Dedicated sheets for summarizing query-level and calculation errors.

* **VBA-Enhanced User Experience & Automation:**
  
  * Scripts for common tasks such as sorting large datasets, clearing filters for analysis, and extracting unique player information.

* **Reporting & Visualization:**
  
  * Designed for results to be published via **Power BI**, with a described data pre-load mechanism using OneDrive to ensure secure and wider accessibility for kingdom members.
