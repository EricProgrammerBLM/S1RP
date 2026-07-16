# SUMMIT Guest Experience Rotation Automator

## Overview
SUMMIT is a sophisticated Python-based scheduling engine designed to automate the complex daily post rotations for the Guest Experience team at Summit One Vanderbilt. It manages multi-floor staff distribution (B1, OB1, OB2, OB3), ensures Tier 1 post coverage, and dynamically adjusts to staffing shortages in real time.

## Key Features
- **Staffing-Aware Logic:** Features "Surgical Lead Logic" that automatically converts Lead float positions into active posts if regular GSA counts fall below the mandatory floor minimum.
- **Contingency Management:** Implements automated "Borrow from B1" and "Freight Sacrifice" protocols to maintain upstairs operational standards during low-attendance shifts.
- **Tiered Priority System:** Posts are classified into Tiers 1-3, ensuring that critical positions (like Rise Tablet or Unity) are never left vacant.
- **Smart Break Scheduling:** Distributes breaks across three windows (6:30, 7:30, 8:30) while ensuring that floor leads do not have conflicting break times.
- **Elevator (CAB) Priority:** Automatically assigns Full-Timers (3:30 PM starters) and Leads to high-priority CAB positions with back-to-back prevention.
- **Mega-Merge Output:** Post-processing logic merges adjacent identical cells in the final Excel output for a clean, professional aesthetic.

## Advanced Logic Modules
- **Check & Balance (Relief Check):** A recursive algorithm that validates every rotation link to ensure every employee is properly relieved by an incoming staffer or breaker.
- **Floor Reversing:** Randomly swaps post-creation directions between OB1 and OB2 for specific columns to diversify staff experience.
- **Natural Sorting:** Uses `natsort` to handle alphanumeric Excel coordinates without logical breaks.

## Requirements
- `openpyxl`
- `xlsxwriter`
- `natsort`

## Usage
1. Open `SUMMIT.py` and update the `File_Name` path to your local environment.
2. Set the `Full_Time_Closers`, `Closers`, `Leads`, and `Breakers` variables based on that day's roster.
3. Run the script:
   ```bash
   python SUMMIT.py
