ARENA operating data - manual workflows
======================================

Workbook: ARENA_OperatingData_ManualWorkflows.xlsm
Sources: ARENATransientDataFormat.xlsx and ARENAInputDataTest.xlsx

This is a separate, self-contained workbook. The original source workbooks are
not modified. Edit or paste the latest input data into the corresponding tabs
in this macro-enabled workbook. It does not automatically reload the originals.
It can be used in any folder. No macros run on open or when a cell changes.

How to run
----------
1. Open the .xlsm workbook in desktop Excel and enable macros.
2. Update FL6 data for steady state and FL6_WD!A:B for the depth survey.
3. Update the operation/life transient input tabs, or add new ones as below.
4. On Workflow Controls, select the input time row in C17. The default is 2.
5. Click Run All Operating Data, or an individual workflow button:
   - 1. Update Steady State
   - 2. Update Temperature Transients
   - 3. Update Pressure Transients
   - 4. Generate PCD/HCD Startup Profiles
   - 5. Define Input Data Range (saves a selection; does not run calculations)
6. Read Workflow Log, resolve missing inputs or warnings, rerun, and save.
   Each run replaces the log with messages for that run. If a stage fails,
   earlier completed stages may already have updated their output.

Transient input layout
----------------------
- A1: optional source description.
- Row 2: time values (default time source).
- Row 3: transient/profile names, copied exactly to output row 2.
- Row 4: source unit headings.
- Row 5: other input metadata; optionally selected as the time source in C17.
- Row 6 onward: A = KP/distance in feet; B onward = raw profile values.
- The raw block ends at the first completely empty column after A, from row 2
  through the last distance row. Do not leave a gap inside the raw block.
- Temperature inputs are Fahrenheit. Pressure inputs are psi. The conversion
  does not change the absolute/gauge basis supplied by the user.
- Maintain a blank separator before the conversion area. Insert columns when
  enlarging the raw block so that raw data never overwrites the conversion area.
- Clear old raw rows and profile columns when replacing data with a smaller set.
  The macro detects the new extent and clears obsolete generated results.
- Extra scratch data after the first blank column is not part of the raw block.
  The conversion block is identified by its Distance heading on row 4.

Transient case names
--------------------
Examples: FCD_CD_EarlyTemp, FCD_HU_EarlyPressure, HCD_HU_MiddleTemp,
PCD_CD_LatePressure. Temperature may end in Temp or Temperature; pressure ends
in Pressure. Keep _CD_ or _HU_ in the name.
FCD = full cooldown; HCD = half cooldown; PCD = partial cooldown.
CD = cooldown; HU = startup/heat-up. Use Early, Middle or Late for life.

To add a missing case, copy an existing input tab, rename it, replace the raw
data and metadata, clear obsolete inputs, and rerun. The new case is detected
automatically. No VBA modification is needed. Do not leave copied sample values
in a new life case and treat them as real input.

Transient conversion and output
-------------------------------
- Distance: feet x 0.3048 = meters.
- Temperature: (Fahrenheit - 32) x 5/9 = Celsius.
- Pressure: psi x 0.068948 = bar, retaining the template conversion factor.
- Every raw profile column is converted in order, including column C. Some
  formulas in the original sample skipped C; the new workflow does not.
- Converted formulas are rebuilt to the current raw row/column extent.
- PressureTransient and TemperatureTransient use the original Early output
  layouts, renamed to accommodate all supplied lives and operation cases.
- Row 1: one merged amplitude heading per case, obtained by removing the
  trailing Temp/Temperature/Pressure suffix from its input tab name.
- Row 2: exact profile labels from raw input row 3.
- Row 3: exact values from the time row selected in C17. Values are not sorted,
  normalized, converted to seconds, or otherwise changed.
- Rows 4 and 5: quantity and units, retaining the template header layout.
- A6: KP=0 boundary when the first input distance is positive. Boundary values
  equal the first profile values, as in the template. If input already starts
  at KP=0, no duplicate zero row is inserted.
- The remaining rows copy the converted KP and profile values, in input order.
- All cases of one quantity must use the same increasing KP grid and row count.
  The macro stops with an explanatory message if grids differ; it does not
  silently resample or misalign profiles. All cases are validated before that
  quantity's existing combined output is replaced.
- The output grows or shrinks in rows and columns, retaining the header layout
  and template formatting. Obsolete generated values and amplitude merges are
  removed when cases/profiles/rows are removed.
- Combined outputs are snapshots of the most recent manual run. Rerun after
  editing raw inputs, times, or labels.

Steady-state workflow
---------------------
- Preserves the supplied FL6 data column groups and their formula conventions.
- Uses the last populated raw distance row in column C, starting at row 6, to
  size Temperature and Pressure, including their initial KP=0 row.
- Column C is treated as feet, following the actual template values and
  conversion formulas. Its original source heading says miles; do not multiply
  those existing C values by 5280 again. I and O retain H x 5280 and N x 5280.
- Rebuilds the temperature/pressure conversion formulas and clears obsolete rows.
- Rebuilds Density min/max/average formulas using each source column's last row.
- FL6_WD A:B are raw distance/depth in feet, starting at row 2. Depth is negative
  in the supplied convention. Its conversion, linear interpolation and pressure
  formulas are resized to actual input and output extents.
- Water-depth interpolation is limited to the supplied survey coverage. Missing
  coverage shows #N/A, rather than extrapolating toward an empty zero-valued row.
- Landing-temperature interpolation retains the existing Y11:Z15 control points
  on Temperature. Reverse KP now refers to the current final steady-state row.
- Source values must be numeric; missing profile inputs show #N/A rather than
  being converted into plausible zero pressure or -17.78 C temperature.
- Late FCD output slots are Temperature column R and Pressure column V.
  When real late FCD data are available in FL6 data, enter their source column
  letters in Workflow Controls C35 (temperature) and C36 (pressure), and their
  common first data row in C37 (default 8). Source rows must align with the
  steady-state KP rows. Blank source settings leave late FCD explicitly missing.

Known limitations of the supplied sample
---------------------------------------
- Eighteen transient input tabs are supplied: six early-life cases and three
  middle-life startup cases, for both temperature and pressure.
- Middle-life cooldown and all late-life transient tabs are absent. The log
  lists missing cases only within a life for which transient data were supplied.
  Middle/late lives without inputs are skipped. No life is inferred from another.
- Some cooldown time sequences in row 2 are not increasing. The workflow copies
  them exactly and logs a warning. Confirm/correct them before using ARENA input.
- No late-life FCD steady-state profile is supplied.
- Early/middle FCD source data in BN/BK finish two rows short of the main steady
  grid under the existing template row mapping; those missing values show #N/A.
- The template's middle FCD columns copy early FCD (Temperature Q=P, Pressure
  U=T). This assumption is preserved and reported in the log.
- The template Density headings called FCD refer to HCD sources AP/AW/BD. This
  original mapping is preserved and reported; confirm its intended meaning.
- The sample water-depth survey ends at 1185 ft, while the steady-state route
  ends at 7838.61 ft. Extend FL6_WD A:B before using depth-dependent pressure
  results beyond the survey coverage.

Generating missing PCD/HCD startup profiles
-----------------------------------------
Button 4 generates missing HCD_HU and PCD_HU cases and refreshes both transient
outputs. Run All also performs this step. Nothing runs automatically on input
changes or workbook open.

The generation is independent for temperature and pressure, and for each life:
- Existing user-supplied HCD_HU/PCD_HU inputs take priority and are retained.
- A missing or empty input can be generated only if matching-life FCD_HU and
  the required steady-state data are supplied.
- PCL (6-hour) in FL6 data supplies the PCD starting profile.
- HCL (12-hour) in FL6 data supplies the HCD starting profile.
- The first generated line is the complete steady-state profile at time zero.
  If the FCD input already has time zero, that line is replaced by the starting
  profile. Otherwise a zero-time line is inserted before all supplied FCD lines.
- At every KP, each subsequent line is MAX(steady-state value, FCD_HU value).
  This is applied to each temperature or pressure profile independently.
- All later FCD time columns and their names/times are retained in order.
- Steady-state values are linearly interpolated by KP where needed. All FCD KPs
  must lie within the supplied steady-state coverage; no extrapolation is used.
- The raw-unit calculation uses feet and Fahrenheit/psi, then the existing
  conversion workflow converts generated results to meters, Celsius and bar.
- Generated inputs are saved in HCD_HU/PCD_HU case tabs and marked Generated in
  Data Ranges column C. Rerunning button 4 refreshes these generated tabs and
  clears obsolete rows/columns. It does not overwrite supplied User inputs.
- If you replace a generated tab with your own measurements, set its origin in
  Data Ranges column C to User before running the generator again.
- If generation becomes unavailable, its tab is marked Generated - unavailable
  and excluded from combined output. Its old data are not treated as current.

FL6 data source columns (KP / temperature / pressure):
  PCD Early:  U / V / W       PCD Middle: Z / AA / AB    PCD Late: AE / AF / AG
  HCD Early:  AK / AL; AM / AN for pressure
  HCD Middle: AR / AS; AT / AU for pressure
  HCD Late:   AY / AZ; BA / BB for pressure
Data begin at row 6, and the selected/available last row is used.

Defining input data ranges
-------------------------
Click button 5, then select the raw rectangle on the input sheet while the range
selection dialog is open. Confirm with OK. The selection is recorded on Data
Ranges, and applies to subsequent manual runs.
- Transient input: start at A6; include distance and the desired profile columns.
  Example: A6:S92 processes rows 6:92 and profiles B:S.
- FL6 data: start at C6. The selected last row limits all its steady-state source
  groups; their existing column mappings remain in use.
- FL6_WD: start at A2 and include B. The last selected row limits the survey.
- Keep the original first-data-row/header layout. Noncontiguous selections and
  transient selections crossing the blank separator are rejected.
- To restore automatic extent detection, enter AUTO in Data Ranges column B.
- You may also edit an existing range address directly in column B.
- For combined transient outputs, every included case of the same quantity must
  use the same KP grid. Apply compatible ranges to those input tabs.
- Range selection does not delete raw inputs outside the selected range.
- Generated tab ranges reset to AUTO so they follow the selected FCD range.

For the supplied sample, steady-state data cover only the first 7838.61 ft,
whereas transient inputs extend farther. Existing supplied startup cases can
still be processed as supplied. To test generation for missing cases, select
compatible shorter transient ranges or supply steady-state data covering the
full FCD KP range.

Optional middle and late life
----------------------------
Only supplied transient lives are processed. Missing middle/late life is not an
error and does not trigger requests to create all of that life's cases.
Steady-state middle/late outputs are also left blank when none of the relevant
temperature, pressure or density inputs for that life is supplied. Their existing
column headings remain. Generated profiles never substitute one life for another.

Verification
------------
The workflow is tested in desktop Excel against raw temperature/pressure values,
profile names, amplitude names, time values, final KP, and the last steady row.
An independent saved-file comparison verifies all 64,182 temperature values and
64,182 pressure values (114 profiles x 563 raw rows for each quantity).
Resize and changed-input checks use a disposable copy, so delivered inputs remain
the source sample. Review Workflow Log before treating any output as ready to use.
The startup fallback was additionally tested against 5,916 temperature/pressure
cutoff values and every generated starting line. Tests also verified supplied
input precedence, selected-range shrinking, and skipping absent middle/late life.
