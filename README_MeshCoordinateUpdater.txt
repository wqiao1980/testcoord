ARENA mesh coordinate updater
=============================

Workbook:
  ARENA_MODELFL6DBMnew3_ManualWorkflows.xlsm

Target worksheet:
  ARENA_MOD_PLWithDBM2-10

How it works
------------
1. Update the mesh anchors in A:D or G:J.
   - A or G = node number
   - B or H = X coordinate
   - C or I = Y coordinate
   - D or J = Z coordinate
   Complete G:J rows take priority; A:D is used as the fallback.

2. Define curves in the contiguous table beginning at row 73.
   - A = start node
   - B = end node
   - D = curve-center X
   - E = curve-center Y
   - I = direction: -1 decreases the angle; 1 increases the angle

3. Nothing runs automatically. Open the Workflow Controls tab and click one
   of the five buttons. Every button first rebuilds columns Z:AC. Z contains
   every node number from the current minimum mesh node through the current
   maximum mesh node; AA, AB, and AC contain X, Y, and Z. Column T is not used
   to determine the output size.

   A larger mesh automatically extends Z:AC. A smaller mesh shortens the
   output and clears obsolete values below the new last node.

4. After Z:AC is rebuilt, its values are copied to Spring2-10!B:E beginning
   at row 5. Every column from G:Y is updated independently from that same
   column's row-6 seed. Formula seeds are filled with relative references;
   separator/value seeds are copied as values. Obsolete rows are cleared when
   the mesh becomes smaller.

5. The Spring2 workflow exports Spring2-10 column Y to this file in the
   same folder as the workbook:
   BPTiber_With_BM_FL6_TestBase_05_Lateral_Springs.txt
   Every literal hzw marker becomes a real line break. The existing lateral-
   spring ELSET footer is preserved and its generated element range is updated.

6. Z:AC is also copied to SpringDistance2-10!G:J beginning at row 9.
   F uses its row-9 starting value and row-10 cumulative formula; K and N:Y
   use their respective row-9 formula or value seeds. F, K, and every column
   N:Y are resized to the same last row as G:J. L:M remains untouched.

7. The SpringDistance workflow exports N:O:P to this file in the same folder
   as the workbook:
   BPTiber_With_BM_FL6_TestBase_05_Initial_1P5m.txt
   The N8 header is retained, data starts at row 9, fields are separated by
   tabs, and column P is written with three decimal places to match the
   workbook display format.

8. Z:AC is also copied to Sheet10!C:F beginning at row 9. The A:O output
   table is resized dynamically: row-9 boundary values are retained, repeating
   formulas use the worksheet's row-10 seeds, and the repeating J formula uses
   its row-11 seed because J10 is a special first-row formula. Sheet10!H5 is
   updated to point to the last populated cell in column A. The separate Q:S
   and CA:CE reference blocks remain untouched.

9. The Sheet10 workflow exports M:N:O to this file in the same folder as the
   workbook:
   BPTiber_With_BM_FL6_TestBase_05_TensionAdj.txt
   M9 supplies the header. The special node-1 line retains its existing field
   spacing and updates only its final value from negative I3, rounded to a
   whole number. Rows 10 through the dynamic last row use M and N exactly as
   stored and write O with three decimal places.

10. Workflow Controls buttons:
    - Run All Workflows
    - 1. Rebuild Coordinates Only
    - 2. Update Spring2 + Lateral TXT
    - 3. Update SpringDistance + Initial TXT
    - 4. Update Sheet10 + Tension TXT

    Each downstream button rebuilds Z:AC first, so it can be run safely by
    itself. The Run All button performs the full sequence.

Notes
-----
- Enable macros when Excel opens the .xlsm file.
- Workbook_Open and Worksheet_Change automation are not present.
- Workflows run only from a button or from Alt+F8.
- Text output paths use ThisWorkbook.Path. You can move the workbook to any
  folder; save it there before running a text-export workflow.
- Straight segments use linear interpolation between consecutive anchors.
- Curve segments use the specified center and direction, with Z interpolated
  between the curve's start and end anchors.
- The original .xlsx file was not overwritten.
