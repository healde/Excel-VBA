## Productivity time tracker

<br>

### 🦕 Purpose
In a supply chain, you may want to record the time per product. When time extends over one day or more, we need to remove every time-off. 
In a context of automation, we want to avoid stopping and starting for each daily breaks. The goal is therefore to subtract them from a single overall time and obtain the most accurate total possible.

#### File shape
One table with some columns for entries and formulas. I suggest to keep a splited formula through columns to maintain clarity. Vba can be used in complement to fill entries, but isn't required by the table for calculation.

<br>

### 🌿 Parameters
Produce Id, Operator, start time, end time   `>`   settings, hidden steps   `>`   first filtered time

#### Process preferences
ProduceId basically remains the record index and can be changed to TaskId, or SectorId, for examples.
