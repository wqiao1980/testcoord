# testcoord

## ARENA operating-data manual workflows

Open [ARENA_OperatingData_ManualWorkflows.xlsm](ARENA_OperatingData_ManualWorkflows.xlsm) in desktop Excel and enable macros. All workflows run manually from the **Workflow Controls** tab.

- Generate missing PCD/HCD startup profiles from matching-life FCD startup profiles and PCL/HCL steady-state inputs, preserving user-supplied startup profiles.
- Use **5. Define Input Data Range** before processing to select the raw data extent; use **4. Generate PCD/HCD Startup Profiles** to generate and update transient outputs.
- Skip middle/late-life processing when those inputs are absent.

See the [operating instructions and range-selection examples](README_ARENA_OperatingData_ManualWorkflows.txt) for required layouts, coverage limits, and verification details. The editable VBA source is [OperatingDataWorkflows.bas](OperatingDataWorkflows.bas).
