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
   of the nine buttons. Numbered buttons 1 through 6 first rebuild columns
   Z:AC. Z contains
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

10. The AsLaid workflow updates AsLaid10DBM2 from ARENA_MOD_PLWithDBM2-10:
    - G:J from row 1 through the row containing the maximum node in column A
      is copied to S:V beginning at row 28.
    - L:S through its dynamic last populated row is copied to AV:BC beginning
      at row 21.
    - Curve rows A73:I75 are copied to AY13:BH15, with BB retained as the
      intentional blank spacer between source columns C and D.
    - W:AG is resized to the node-copy last row and filled from each column's
      row-28 formula or value seed. Obsolete rows are cleared.

11. The AsLaid-to-Sheet10 workflow copies AsLaid10DBM2!B:E beginning at row
    23 to Sheet10!C:F beginning at row 9. Its dynamic last row follows the
    last populated source row across B:E. Every column G:O is then resized and
    filled independently from that column's row-10 formula or value seed.
    Sheet10!H5 is refreshed to reference the actual last populated row in
    column A before the G:O formulas are calculated.
    It exports Sheet10!M:O beginning at row 10 to:
    BPTiber_With_BM_FL6_TestBase_05_TensionAdj.txt
    The existing first two text-file lines are kept exactly; when no prior file
    exists, the two required default lines are created. Data is replaced from
    line 3 onward, with column O written to three decimal places.

12. The ModStructure endpoint workflow copies the current route node, X, and Y
    values from ARENA_MOD_PLWithDBM2-10!A:C to ModStructure!B:D. The source
    ends at the last complete route row before the curve-input table at row 73.
    It then rebuilds the two end structures from the formulas represented on
    ModStructure:
    - G1 supplies the extension distance; its default value is 100 m.
    - N1 supplies the structure height; its default value is 0.75 m.
    - The start structure is projected backward from the first route segment.
    - The end structure is projected forward from the last route segment.
    - Both generated structure node numbers equal their route anchor node plus
      7,000,000.
    The route section expands or shrinks with the source data. Obsolete rows are
    cleared, and the end-structure labels and elevations move to the new last
    row when the route has fewer nodes.

13. The intermediate-structure workflow uses the input block in
    ModStructure!N13:S22. Enter an exact existing route node in O14, then click
    the intermediate-structure button. The selected node must have both a
    preceding and a following route node, so it cannot be the first or last
    route node. The workflow:
    - projects one point upstream and one point downstream by the distance in
      O15, which references G1 and defaults to 100 m;
    - looks up the selected route node's Z coordinate from the generated route
      output first, with the mesh input tables as a fallback;
    - adds the height in O16, which references N1 and defaults to 0.75 m; and
    - writes the two result rows to O21:S22.
    O17 and O18 are editable node-number additions. Their defaults are
    8,000,000 for the upstream structure node and 9,000,000 for the downstream
    structure node.

14. Workflow Controls buttons:
    - Run All Workflows
    - 1. Rebuild Coordinates Only
    - 2. Update Spring2 + Lateral TXT
    - 3. Update SpringDistance + Initial TXT
    - 4. Update Sheet10 + Tension TXT
    - 5. Update AsLaid10DBM2
    - 6. AsLaid to Sheet10 + Tension TXT
    - 7. Update ModStructure End Structures
    - 8. Build Intermediate Structure

    Numbered buttons 1 through 6 rebuild Z:AC first, so each can be run safely
    by itself. Buttons 7 and 8 refresh the ModStructure route directly from
    ARENA_MOD_PLWithDBM2-10!A:C. If mesh elevations have changed, run button 1
    or Run All before button 8 so its Z-coordinate lookup uses current output.
    Run All performs the complete standard sequence, including the ModStructure
    endpoint workflow. It does not run the intermediate-structure workflow
    because that calculation requires a user-selected node in O14.

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
- The ModStructure sheet is self-contained and does not retain external links
  to routeforcode2.xlsx.
- The original .xlsx file was not overwritten.
