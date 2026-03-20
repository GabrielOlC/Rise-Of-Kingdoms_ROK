# Game Data Analytics Engine (Rise of Kingdoms)

**An enterprise-grade ETL pipeline and KPI normalization engine built to govern a 300+ player gaming community through objective data.**

> **Status:** Legacy / Production (2025)
> **Role:** Lead Data Engineer & Kingdom Council Member  
> **Key Result:** Fully automated the ingestion, cleaning, and complex scoring of 300-600 users across multi-week campaign events.
> **Live Dashboard:** [PowerBI Dashboard]([Microsoft Power BI](https://app.powerbi.com/view?r=eyJrIjoiMTU2NmU5ODItNjU0YS00MTEwLTgxODAtMjcxOTJlZGFiMTdlIiwidCI6IjYyZTAxNWZhLTM4MjUtNDRiOC04Yzk5LWVjODhlMjEzYTk3OSJ9)) 

<div align="center">
  <img src="https://img.shields.io/badge/Power_BI-F2C811?style=for-the-badge"/>
  <img src="https://img.shields.io/badge/Power_Query-F2C811?style=for-the-badge"/>
  <img src="https://img.shields.io/badge/Excel-217346?style=for-the-badge"/>
  <img src="https://img.shields.io/badge/VBA-gray?style=for-the-badge"/>

</div>

---

## 🌍 The Origin Story

Managing a "Kingdom" of 300–600 players in a highly competitive MMO (_Rise of Kingdoms_) is equivalent to running a mid-sized company. Leadership must track performance across massive server-vs-server wars (KvK), distribute rewards fairly, and penalize non-contributors—all while maintaining trust in governance.

Originally, this was done manually: disjointed scans, dirty data, and 2–3 hours of spreadsheet math per event. Calculating fair "Contribution Points" (CP) was exhausting and often led to disputes and accusations of favoritism.

I architected a **comprehensive Data Analytics Engine**—a centralized infrastructure that transitioned the kingdom from **subjective governance to absolute data-driven meritocracy.**

This engine automated the entire lifecycle:

* **Raw data extraction** from multiple metric scans

* **Clean processing & score calculations** with Power Query, advanced Excel formulas, and VBA

* **Transparent dashboards** for player evaluations and reward distribution

The result: governance became unbiased, efficient, and scalable, removing human bias and empowering the kingdom to thrive in competitive KvK battles.

---

## 💎 Key Modules & Engineering

### 1. The KPI Normalization Engine (Math & Logic)

 *Solving Performance Evaluation Across Asymmetric Users*

---

**Problem Statement:** "Fairness" in player evaluation is mathematically complex. A top-tier _Whale_ account achieving 10,000 kills is expected; a mid-tier player reaching the same milestone is extraordinary. Raw kill counts alone cannot capture true effort. Additional distortions made the problem even harder:

* **Farm accounts** inflated statistics without reflecting real contribution.

* **Zeroing events** (players being wiped outside official war times) unfairly skewed death counts.

* **Account scale differences** meant whales and standard players could not be compared on the same baseline.

<details>
  <summary><b>🎯 Key Challenges Addressed</b> <i>(Click to expand)</i></summary>
  <ul>
    <li><b>Asymmetric Baselines:</b> How to fairly rank players of vastly different power levels?</li>
    <li><b>Variable Weighting:</b> How to assign different values to different actions (e.g., Tier 4 vs. Tier 5 troop kills/dead)?</li>
    <li><b>Meritocratic Tiers:</b> How to automatically assign players to performance brackets (Green/Top for rewards, Yellow for probation/school, Red for banishment)?</li>
  </ul>
</details>

**The Solution:**
I engineered an `Achievements Contribution` algorithm using advanced Excel Array functions ([Check documentation]([Game-Data-Analytics-Engine/Documentation.md at main · GabrielOlC/Game-Data-Analytics-Engine](https://github.com/GabrielOlC/Game-Data-Analytics-Engine/blob/main/Documentation.md#achievements-contributio))).

* **Proportional Scaling:** The algorithm calculates expected performance based on a player's base Power, dynamically weighting T4 and T5 kills/deaths into a single normalized KPI.
* **Universal Fairness:** By converting raw stats into a "Percentage of Goal Achieved," the system removed human bias, establishing undeniable parameters for rewards or kingdom exile.
  
  

**⚡Impact⚡**

Eradicated subjective council disputes and accusations of favoritism by establishing a universally trusted, math-based meritocracy. Successfully standardized the kingdom's governance into clear, undeniable performance tiers (Reward, Probation, Exile), ensuring 100% fair and transparent evaluations across hundreds of players.



### 2. Temporal ETL & Anomaly Detection (Power Query)

*Solving Messy OCR/Scan Data*

---

**Problem Statement:** Raw data came from third-party Optical Character Recognition (OCR) scans of the game UI, resulting in highly erratic data files filled with nulls, mismatched types, and players who occasionally dropped off the radar, making linear tracking impossible.

<details>
  <summary><b>🎯 Key Challenges Addressed</b> <i>(Click to expand)</i></summary>
  <ul>
    <li><b>Chronological Merging:</b> How to calculate exact metric changes between Event A and Event B?</li>
    <li><b>Data Quarantine:</b> How to catch impossible data (e.g., a player having fewer kills after a war than before)?</li>
    <li><b>Schema Volatility:</b> How to process unstructured folder dumps into a unified relational model?</li>    
    <li><b>Data Cleaning:</b> How to automatically handle null values and mismatched data types without crashing the model?</li>
  </ul>
</details>

**The Solution:**
A robust data ingestion pipeline built in **M (Power Query)** and **VBA**.

* **Automated Delta Calculations:** Ingests specifically named block files (e.g., `F1.A` and `F1.B`), joins them on `Governor ID`, and calculates the exact resource delta for that specific time period.
* **The Quarantine Layer:** Built `Check Data Structure` queries to automatically flag Type Mismatches, Negative Values, and missing IDs, forcing data cleaning *before* it hits the calculation engine.
* **Scalability:** Processed up to 1,000,000 rows of data, accepting any kingdom size.
  
  

**⚡Impact⚡**

Data analysis becomes almost fully autonomous. Updating the model with new event data takes less than one minute. Dashboards highlight common errors and provide mechanisms to fix them quickly. Contribution Points (CP), player Achievements Contribution, and ranks are all calculated automatically based on defined parameters.



### 3. Virtual Economy & Relational Mapping

*Solving Currency Distribution and Multi-Account Tracking*

---

**Problem Statement:** Players utilized secondary "Farm" or "Alt" accounts, making it difficult to track a single human's total contribution. Furthermore, we needed a transparent ledger for the kingdom's virtual currency (Contribution Points - CP).

<details>
  <summary><b>🎯 Key Challenges Addressed</b> <i>(Click to expand)</i></summary>
  <ul>
    <li><b>Entity Resolution:</b> How to map multiple child accounts to a single parent account?</li>
    <li><b>Economy Ledger:</b> How to track CP earned vs. CP spent on in-game auctions?</li>
    <li><b>Economy time-line:</b> How to remove earned CP not used from old KVKs to avoid low contributors to save currency to high ranks?</li>
  </ul>
</details>

**The Solution:**

* **Parent/Child Rollup:** Implemented relational mapping allowing Contribution Points generated by "Alt" accounts to be automatically summed and credited to the "Main" account.
* **Public Ledger:** The system calculated total available CP, subtracting spent points and CP not used from OLD kvks, creating an indisputable "bank account" for every player to bid on rewards and ranks.
  
  

**⚡Impact⚡**

Players can view their currency on an event‑by‑event basis, allowing them to align personal goals with the realities of KvK. This transparency increases competitiveness, motivating players to earn more Contribution Points (CP) by actively engaging in battles.



### 4. Secure Data Publishing Architecture

*Solving IP Protection & Public Transparency*

---

**Problem Statement:** We needed to share the results publicly so players could track their money (CP) and ranks, but sharing the raw Excel files risked exposing the algorithms or allowing data tampering.

**The Solution:**

* **Decoupled Architecture:** Created a "Pre-Load" M-Query that strips all formulas and logic, pushing only the flattened results to a secondary "White Access" file via OneDrive.
* **Power BI Integration:** Connected the live Power BI dashboard strictly to the flattened White Access file. This ensured 100% public transparency while maintaining zero security risk to the source code.
  
  

---

## ⚡ Strategic Impact

* **Diplomatic Auditing:** The robust nature of our system allowed our Kingdom to independently audit the data provided by allied kingdoms during cross-server events, identifying discrepancies and ensuring fair coalition scoring.

---

## 📂 Repository Structure

```text
/Game-Data-Analytics-Engine (preview)
├── READme
├── Documentation    # Doc with the explanation of the code and the updates
|
├── /1_ETL_and_Data_Cleaning        # Power Query (M) scripts for data ingestion and quarantine
├── /2_Algorithms_and_Logic         # Complex KPI normalization and CP calculation formulas
└── /assets                         # 
```
