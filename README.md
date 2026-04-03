# TZJ501-M484-3375-SEQ

## COATER Column Configuration

The 'slot' order is important for both TEST and POR data and corresponds to a COATER. A COATER column has been created for each data set.

### TEST Data COATER Mapping
- **8 COATER positions** based on wafer order within each XML file:
  - **Pattern**: Each XML file starts fresh with wafer #1
  - 1st wafer in file → COATER '1301'
  - 2nd wafer in file → COATER '1401'  
  - 3rd wafer in file → COATER '1303'
  - 4th wafer in file → COATER '1403'
  - 5th wafer in file → COATER '1302'
  - 6th wafer in file → COATER '1402'
  - 7th wafer in file → COATER '1304'
  - 8th wafer in file → COATER '1404'
  - 9th wafer in file → COATER '1301' (pattern repeats)

### POR Data COATER Mapping
- **3 slots per run** (each run = unique LOT + LOT_DATA_COLLECT_DATE combo)
- Slot ordering within each run:
  - 1st slot → COATER '1'
  - 2nd slot → COATER '2' 
  - 3rd slot → COATER '3'

## How to Update COATER Values

### Method 1: Edit Configuration Variables (Recommended)
The COATER mappings are defined at the top of each Python file for easy modification:

**In `dash_app.py` and `SEQ_dash_app.py`:**
```python
# TEST Data: 8 COATER positions based on wafer order within each XML file
TEST_COATER_BASE_PATTERN = ['1301', '1401', '1303', '1403', '1302', '1402', '1304', '1404']

# COATER assignment logic (automatically applied):
# 1st wafer in XML file -> COATER 1301
# 2nd wafer in XML file -> COATER 1401
# 3rd wafer in XML file -> COATER 1303
# ... and so on, cycling through the pattern

# DECK Mapping: Group coaters into decks
COATER_TO_DECK_MAPPING = {
    '1301': '13-L', '1302': '13-L',  # Deck 13 Left
    '1303': '13-R', '1304': '13-R',  # Deck 13 Right 
    '1401': '14-L', '1402': '14-L',  # Deck 14 Left
    '1403': '14-R', '1404': '14-R'   # Deck 14 Right
}

# POR Data: COATER values for slots 1, 2, 3 (easily configurable)
POR_COATER_VALUES = ['1', '2', '3']

# TEST Entity: Hardcoded entity for TEST data (easily configurable)
TEST_ENTITY = 'TZJ501'  # Change this value to update the TEST entity
```

**To change values:**
1. Open the Python file(s)
2. Modify the `TEST_COATER_BASE_PATTERN` list to change coater values
3. Modify other configuration variables in the sections at the top (lines 15-35)
4. Save the file

**Examples of alternative configurations:**
```python
# Example: Using different TEST coater pattern (still 8 values, cycles through slots 1-25)
TEST_COATER_BASE_PATTERN = ['2001', '2002', '2003', '2004', '2005', '2006', '2007', '2008']

# Example: Using letter codes for POR
POR_COATER_VALUES = ['A', 'B', 'C']

# Example: Using descriptive names for POR  
POR_COATER_VALUES = ['COAT1', 'COAT2', 'COAT3']

# Example: Changing TEST entity
TEST_ENTITY = 'TTG624'  # Or any other entity code
```

### Method 2: Configuration File (Future Enhancement)
A `coater_config.ini` file has been created for reference showing how configuration could be externalized in the future.

**Note:** All COATER values are stored as text categories, not integers.