# Supply Chain Cost & Logistics Performance Analysis

Analisis biaya dan kinerja operasional outbound logistics menggunakan Microsoft Excel, Power Query, formula Excel, PivotTable, dan interactive dashboard.

Project ini menggunakan 9.215 historical orders untuk menganalisis struktur biaya, kapasitas warehouse, freight rate matching, serta menemukan penyebab transaksi yang tidak memiliki freight rate.

> **Project Scope:** Descriptive cost and capacity analysis. Project ini tidak mencakup network optimization atau Linear Programming.

---

## Project Overview

Project ini dibuat untuk menjawab beberapa pertanyaan utama:

- Berapa besar biaya warehouse dan freight terhadap total operational cost?
- Plant mana yang memiliki volume order paling tinggi dibandingkan kapasitas hariannya?
- Bagaimana biaya freight dihitung berdasarkan weight, freight rate, dan minimum charge?
- Mengapa sebagian order tidak memiliki freight rate yang cocok?
- Apakah terdapat pola tertentu pada unmatched freight rate?
- Bagaimana hasil analisis dapat disajikan dalam bentuk report dan dashboard?

Analisis dilakukan dari data historis menggunakan Microsoft Excel tanpa melakukan network optimization.

---

## Dataset

**Source:** Supply Chain Logistics Problem Dataset — Brunel University London

Dataset terdiri dari 7 tabel utama:

- `OrderList`
- `FreightRates`
- `WhCosts`
- `WhCapacities`
- `ProductsPerPlant`
- `PlantPorts`
- `VmiCustomers`

Dataset berisi **9.215 historical orders**, dengan informasi mengenai order, plant, port, carrier, service level, transit time, weight, warehouse cost, dan freight rate.

### Scope Dataset

Dokumentasi dataset menyediakan dua kemungkinan penggunaan:

1. Menghitung biaya dari historical network.
2. Melakukan network optimization menggunakan Linear Programming.

Project ini hanya berfokus pada **penggunaan pertama**, yaitu descriptive analysis terhadap biaya dan kapasitas.

---

## Tools & Techniques

### Microsoft Excel

Digunakan untuk:

- Data modeling
- Business calculations
- Cost calculations
- Lookup dan multi-condition matching
- PivotTable dan PivotChart
- Dashboard
- Conditional Formatting

### Power Query

Digunakan untuk:

- Import data
- Promote headers
- Explicit data type transformation
- Column cleanup
- Renaming columns
- Data preparation sebelum analisis

### Excel Formulas

Beberapa fungsi yang digunakan:

- `XLOOKUP`
- `MINIFS`
- `COUNTIFS`
- `MAX`
- Array-based multi-condition matching

Formula freight cost menggunakan konsep:

`MAX(Weight × Freight Rate, Minimum Freight Cost)`

Sehingga minimum shipment charge tetap diperhitungkan ketika hasil perhitungan berdasarkan weight dan rate lebih rendah dari minimum charge.

---

## Data Preparation

Tujuh source tables diproses secara terpisah menggunakan Power Query.

Beberapa proses utama:

- Promote headers
- Explicit data type transformation
- Remove empty trailing columns
- Rename inconsistent columns
- Standardize key fields

Contoh:

`Plant ID` pada `WhCapacities` diubah menjadi `Plant Code` agar konsisten dengan tabel lainnya.

Untuk detail proses Power Query, lihat:

`documentation/power-query-steps.md`

> **Refresh Note:** Power Query pada workbook menggunakan source file lokal. Jika workbook ingin di-refresh pada komputer lain, lokasi source file perlu disesuaikan kembali.

---

# Analysis Methodology

## 1. Master Data

Seluruh data yang telah dipersiapkan digabungkan menjadi dataset analisis.

`Master Data` berisi informasi seperti:

- Order ID
- Order Date
- Origin Port
- Destination Port
- Carrier
- Transit Time
- Service Level
- Plant
- Quantity
- Weight
- Warehouse Cost
- Freight Rate
- Minimum Freight Cost
- Total Freight Cost

---

## 2. Freight Rate Matching

Freight rate dicocokkan berdasarkan beberapa kondisi:

- Carrier
- Origin Port
- Destination Port
- Service Level
- Transit Time (TPT)
- Weight berdasarkan minimum dan maximum weight band

Jika terdapat beberapa freight rate yang cocok, logic matching memilih rate minimum dan kemudian minimum freight cost yang sesuai.

Jika tidak ditemukan rate yang memenuhi seluruh kondisi, transaksi dikategorikan sebagai:

`Not Found`

---

## 3. Freight Rate Gap Investigation

Untuk mengetahui penyebab `Not Found`, analisis dilakukan secara bertahap:

1. Filter seluruh order dengan freight rate `Not Found`.
2. Analisis distribusi weight dari unmatched orders.
3. Identifikasi carrier, route, service level, dan TPT yang terkait.
4. Periksa weight bands pada tabel `FreightRates`.
5. Bandingkan weight range order dengan weight bands yang tersedia.
6. Identifikasi interval weight yang tidak memiliki coverage.
7. Hitung jumlah order yang berada di masing-masing uncovered interval.

Pendekatan ini digunakan agar weight gap **ditemukan dari data**, bukan ditentukan terlebih dahulu melalui hardcoded threshold.

---

# Key Findings

## 1. Warehouse Cost Is the Main Cost Driver

Warehouse cost mencapai sekitar **99,56%** dari calculated operational cost.

- Warehouse Cost: **$15.633.222**
- Freight Cost: **$69.568**
- Calculated Operational Cost: **$15.702.790**

Hal ini menunjukkan bahwa dalam dataset yang berhasil dihitung, warehouse cost jauh lebih dominan dibandingkan freight cost.

> Catatan: Freight Cost hanya berasal dari order dengan freight rate yang berhasil ditemukan.

---

## 2. Order Volume Is Highly Concentrated in PLANT03

PLANT03 menangani:

- **8.541 dari 9.215 orders**
- sekitar **92,7%** dari seluruh historical orders
- Daily Capacity: **1.013 orders**
- Order-to-Capacity Ratio: **843%**

Tiga plant lainnya juga memiliki order volume di atas stated daily capacity:

- PLANT08
- PLANT12
- PLANT09

> Order-to-Capacity Ratio menunjukkan perbandingan volume order terhadap stated daily capacity. Angka ini tidak dapat dianggap sebagai actual utilization karena dataset tidak menyediakan informasi shift, processing time, atau backlog.

---

## 3. Freight Rate Gap Is Concentrated in One Route Combination

Terdapat:

- **1.370 orders**
- sekitar **14,87%** dari seluruh orders non-CRF

yang tidak memiliki matching freight rate.

Seluruh unmatched orders tersebut berasal dari kombinasi:

- Carrier: `V444_1`
- Origin Port: `PORT04`
- Destination Port: `PORT09`
- Service Level: `DTD`

### Weight Band Gap

Investigasi terhadap freight rate table menunjukkan adanya uncovered weight intervals.

Gap terbesar berada antara:

**> 2,50 kg hingga < 70,51 kg**

Dari 1.370 unmatched orders:

- **1.364 orders (99,56%)** berada pada gap besar tersebut.
- **6 orders** berada pada gap weight band kecil lainnya di sekitar batas 0,50–0,51 kg dan 1,00–1,01 kg.

Temuan ini menunjukkan bahwa unmatched freight rate bukan terjadi secara acak, tetapi berkaitan dengan coverage weight bands pada freight rate table.

---

## 4. Minimum Freight Charge Has a Significant Impact

Sebagian besar order dengan valid freight rate memiliki freight cost yang ditentukan oleh minimum shipment charge.

Sekitar **88,97%** dari orders dengan valid rate menggunakan minimum charge sebagai dasar cost.

Karena itu, penggunaan logic:

`MAX(Weight × Rate, Minimum Freight Cost)`

diperlukan agar freight cost calculation mengikuti struktur minimum charge pada rate table.

---

## 5. VMI Transactions Are Structurally Separate

Empat plant tercatat pada `VmiCustomers`, tetapi tidak muncul pada `OrderList`.

Hal ini menunjukkan bahwa transaksi Vendor Managed Inventory pada dataset ini tidak tercatat melalui standard order tracking yang sama.

Temuan ini dicatat sebagai data-structure observation, bukan sebagai kesimpulan mengenai proses operasional perusahaan sebenarnya.

---

# Workbook Structure

| Sheet | Purpose |
|---|---|
| `PQ_WhCosts` | Source table — warehouse costs |
| `PQ_WhCapacities` | Source table — warehouse capacities |
| `PQ_VmiCustomers` | Source table — VMI customers |
| `PQ_ProductsPerPlant` | Source table — products by plant |
| `PQ_PlantPorts` | Source table — plant-port mapping |
| `PQ_OrderList` | Source table — historical orders |
| `PQ_FreightRates` | Source table — freight rate bands |
| `Master Data` | Cleaned and joined analytical dataset |
| `Warehouse Report` | Order volume versus stated daily capacity |
| `Finance Report` | Cost analysis by carrier |
| `Freight Rate Gap Analysiss` | Investigation of unmatched freight rates |
| `Dashboard Data` | Supporting data for dashboard |
| `Dashboard Summary` | Executive dashboard |
| `Documentation` | Business rules, assumptions, methodology, and data-quality findings |

---

# Repository Structure

```text
supply-chain-logistics-analysis/
│
├── README.md
│
├── supply-chain-logistics-analysis.xlsx
│
├── raw-dataset/
│   └── Supply_chain_logistics_problem.xlsx
│
├── asset/
│   ├── dashboard-summary.png
│   ├── warehouse-report.png
│   └── finance-report.png
│
└── documentation/
    └── power-query-steps.md
