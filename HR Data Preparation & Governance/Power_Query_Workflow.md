# Power Query workflow: HR monthly workforce preparation

Build this in one Excel workbook. Each query is pasted into **Home > Advanced Editor**. Step names are written to read clearly in the **Applied Steps** pane.

## 0. Setup (1 min)

1. Put all supplied files in one folder, e.g. `C:\HR_Project\`. Use the original file names.
2. In Excel: **Data > Get Data > Launch Power Query Editor**.
3. **Home > Manage Parameters > New Parameter**
   - Name: `SourceFolder` · Type: **Text** · Current Value: `C:\HR_Project\` (keep the trailing `\`)
   - This is the only place the path appears, so the workflow refreshes on any machine.
4. For each query below: **Home > New Source > Other Sources > Blank Query**, then **Advanced Editor**. Paste the code, then rename the query (right-click > Rename) to the name shown.

## Query map

| Query | Purpose | Load to |
|---|---|---|
| `Src_Monthly` | Raw monthly extract, unchanged | Connection only |
| `Src_Reference` | HR reference extract: trimmed, typed, fields prefixed `Ref_` | Connection only |
| `Src_ApprovedValues` | Approved Values sheet from the rules workbook | Connection only |
| `Recon_IDs` | Step 1: matched / unmatched / repeated EmployeeIDs | Sheet `Reconciliation` |
| `Recon_Field_Compare` | Step 1: visible field differences, monthly vs reference | Sheet `Reconciliation` (below) |
| `Workforce_Validated` | Steps 2–3: cleanse, enrich, validate (all records) | Sheet `Validation` |
| `Workforce_Exceptions` | Step 3: records that failed a rule | Sheet `Exceptions` |
| `Workforce_Final` | Step 4: final reporting population → CSV + PivotTable | Sheet `Final` |

To set the load destination: **Home > Close & Load To…** Choose **Only Create Connection** for the `Src_` queries and **Table > New/Existing worksheet** for the others.

---

## Src_Monthly

```powerquery
let
    Source = Csv.Document(File.Contents(SourceFolder & "Monthly_Workforce_Extract.csv"), [Delimiter = ",", Columns = 9, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true])
in
    #"Promoted Headers"
```

## Src_Reference

```powerquery
let
    Source = Csv.Document(File.Contents(SourceFolder & "HR_Reference_Extract.csv"), [Delimiter = ",", Columns = 6, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
    #"Promoted Headers" = Table.PromoteHeaders(Source, [PromoteAllScalars = true]),
    #"Trimmed Text" = Table.TransformColumns(#"Promoted Headers", List.Transform(Table.ColumnNames(#"Promoted Headers"), each {_, Text.Trim, type text})),
    #"Changed Type" = Table.TransformColumnTypes(#"Trimmed Text", {{"EffectiveDate", type date}}, "en-US"),
    #"Renamed As Reference Fields" = Table.RenameColumns(#"Changed Type", {{"DepartmentName", "Ref_Department"}, {"Region", "Ref_Region"}, {"WorkerStatus", "Ref_Status"}, {"WorkerType", "Ref_Type"}, {"EffectiveDate", "Ref_EffectiveDate"}})
in
    #"Renamed As Reference Fields"
```

## Src_ApprovedValues

Close `HR_Project_Validation_Rules.xlsx` in Excel before refreshing.

```powerquery
let
    Source = Excel.Workbook(File.Contents(SourceFolder & "HR_Project_Validation_Rules.xlsx"), null, true),
    #"Approved Values Sheet" = Source{[Item = "Approved Values", Kind = "Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Approved Values Sheet", [PromoteAllScalars = true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers", {{"Category", type text}, {"Approved Value", type text}})
in
    #"Changed Type"
```

---

## Recon_IDs (Step 1 evidence)

```powerquery
let
    Source = Src_Monthly,
    #"Counted Monthly Rows Per ID" = Table.Group(Source, {"EmployeeID"}, {{"Monthly_Rows", each Table.RowCount(_), Int64.Type}}),
    #"Merged Reference IDs" = Table.NestedJoin(#"Counted Monthly Rows Per ID", {"EmployeeID"}, Src_Reference, {"EmployeeID"}, "Reference", JoinKind.FullOuter),
    #"Expanded Reference ID" = Table.ExpandTableColumn(#"Merged Reference IDs", "Reference", {"EmployeeID"}, {"Reference_EmployeeID"}),
    #"Added Recon EmployeeID" = Table.AddColumn(#"Expanded Reference ID", "Recon_EmployeeID", each if [EmployeeID] <> null then [EmployeeID] else [Reference_EmployeeID], type text),
    #"Replaced Null Row Counts" = Table.ReplaceValue(#"Added Recon EmployeeID", null, 0, Replacer.ReplaceValue, {"Monthly_Rows"}),
    #"Added Match Status" = Table.AddColumn(#"Replaced Null Row Counts", "MatchStatus", each if [Monthly_Rows] > 0 and [Reference_EmployeeID] <> null then "Matched" else if [Monthly_Rows] > 0 then "Monthly only" else "Reference only", type text),
    #"Added Repeated ID Flag" = Table.AddColumn(#"Added Match Status", "RepeatedInMonthly", each if [Monthly_Rows] > 1 then "Yes" else "No", type text),
    #"Selected Recon Columns" = Table.SelectColumns(#"Added Repeated ID Flag", {"Recon_EmployeeID", "Monthly_Rows", "MatchStatus", "RepeatedInMonthly"}),
    #"Sorted By EmployeeID" = Table.Sort(#"Selected Recon Columns", {{"Recon_EmployeeID", Order.Ascending}})
in
    #"Sorted By EmployeeID"
```

Expected result: 12 rows, with 11 `Matched` and 1 `Reference only` (E1012). E1011 has `Monthly_Rows` = 2, `RepeatedInMonthly` = Yes.

## Recon_Field_Compare (Step 1 evidence: visible differences vs confirmed errors)

```powerquery
let
    Source = Table.Distinct(Src_Monthly),
    #"Merged HR Reference" = Table.NestedJoin(Source, {"EmployeeID"}, Src_Reference, {"EmployeeID"}, "Reference", JoinKind.Inner),
    #"Expanded HR Reference" = Table.ExpandTableColumn(#"Merged HR Reference", "Reference", {"Ref_Department", "Ref_Status", "Ref_Type"}),
    #"Added Department Differs" = Table.AddColumn(#"Expanded HR Reference", "Department_Differs", each [Department] <> [Ref_Department], type logical),
    #"Added Status Differs" = Table.AddColumn(#"Added Department Differs", "Status_Differs", each [EmploymentStatus] <> [Ref_Status], type logical),
    #"Added Type Differs" = Table.AddColumn(#"Added Status Differs", "Type_Differs", each [EmploymentType] <> [Ref_Type], type logical),
    #"Kept Rows With Differences" = Table.SelectRows(#"Added Type Differs", each [Department_Differs] or [Status_Differs] or [Type_Differs]),
    #"Selected Compare Columns" = Table.SelectColumns(#"Kept Rows With Differences", {"EmployeeID", "Department", "Ref_Department", "EmploymentStatus", "Ref_Status", "EmploymentType", "Ref_Type", "Department_Differs", "Status_Differs", "Type_Differs"})
in
    #"Selected Compare Columns"
```

Expected result: 5 rows (E1002, E1004, E1007, E1009, E1010). This compares raw monthly values, so differences in case and spacing show up here.

---

## Workforce_Validated (Steps 2–3: main transformation; capture its Applied Steps)

```powerquery
let
    Source = Src_Monthly,
    #"Removed Exact Duplicates" = Table.Distinct(Source),
    #"Trimmed Text" = Table.TransformColumns(#"Removed Exact Duplicates", List.Transform(Table.ColumnNames(#"Removed Exact Duplicates"), each {_, Text.Trim, type text})),
    #"Blanks To Null" = Table.ReplaceValue(#"Trimmed Text", "", null, Replacer.ReplaceValue, Table.ColumnNames(#"Trimmed Text")),
    #"Changed Type" = Table.TransformColumnTypes(#"Blanks To Null", {{"HireDate", type date}, {"TerminationDate", type date}, {"FTE", type number}}, "en-US"),
    #"Standardized Status" = Table.TransformColumns(#"Changed Type", {{"EmploymentStatus", Text.Proper, type text}}),
    #"Standardized Type" = Table.ReplaceValue(#"Standardized Status", "Full Time", "Full-Time", Replacer.ReplaceValue, {"EmploymentType"}),
    #"Merged HR Reference" = Table.NestedJoin(#"Standardized Type", {"EmployeeID"}, Src_Reference, {"EmployeeID"}, "Reference", JoinKind.LeftOuter),
    #"Expanded HR Reference" = Table.ExpandTableColumn(#"Merged HR Reference", "Reference", {"Ref_Department", "Ref_Region"}),
    #"Department From Reference" = Table.ReplaceValue(#"Expanded HR Reference", each [Department], each if [Ref_Department] <> null then [Ref_Department] else [Department], Replacer.ReplaceValue, {"Department"}),
    #"Region From Reference" = Table.RenameColumns(#"Department From Reference", {{"Ref_Region", "Region"}}),
    ApprovedStatus = Table.SelectRows(Src_ApprovedValues, each [Category] = "EmploymentStatus")[Approved Value],
    ApprovedType = Table.SelectRows(Src_ApprovedValues, each [Category] = "EmploymentType")[Approved Value],
    ApprovedDepartment = Table.SelectRows(Src_ApprovedValues, each [Category] = "Department")[Approved Value],
    #"Added Exception Reason" = Table.AddColumn(#"Region From Reference", "ExceptionReason", each Text.Combine(List.RemoveNulls({
            if [EmployeeID] = null then "VR-01 EmployeeID blank" else null,
            if not List.Contains(ApprovedStatus, [EmploymentStatus]) then "VR-02 EmploymentStatus not approved" else null,
            if not List.Contains(ApprovedType, [EmploymentType]) then "VR-03 EmploymentType not approved" else null,
            if not List.Contains(ApprovedDepartment, [Department]) then "VR-04 Department not approved" else null,
            if [HireDate] = null then "HireDate blank" else null,
            if [TerminationDate] <> null and [HireDate] <> null and [TerminationDate] < [HireDate] then "VR-05 TerminationDate before HireDate" else null,
            if [FTE] = null or [FTE] <= 0 or [FTE] > 1 then "VR-08 FTE out of range" else null,
            if [Region] = null then "VR-10 No reference Region" else null
        }), "; "), type text),
    #"Added Validation Status" = Table.AddColumn(#"Added Exception Reason", "ValidationStatus", each if Text.Contains([ExceptionReason], "VR-05") then "Excluded" else if [ExceptionReason] <> "" then "Flagged" else "Valid", type text),
    #"Removed Helper Columns" = Table.RemoveColumns(#"Added Validation Status", {"Ref_Department"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Helper Columns", {"EmployeeID", "EmployeeName", "Department", "Location", "Region", "EmploymentStatus", "EmploymentType", "HireDate", "TerminationDate", "FTE", "ValidationStatus", "ExceptionReason"}),
    #"Sorted By EmployeeID" = Table.Sort(#"Reordered Columns", {{"EmployeeID", Order.Ascending}})
in
    #"Sorted By EmployeeID"
```

Rule mapping for each step (for the PDF):

| Applied step | Rule |
|---|---|
| Removed Exact Duplicates | VR-06 (runs on raw rows, so only true exact duplicates collapse) |
| Trimmed Text, Blanks To Null, Changed Type | VR-09 formatting / data types |
| Standardized Status | VR-02 (`ACTIVE`, `active ` → `Active`) |
| Standardized Type | VR-03 (`Full Time` → `Full-Time`) |
| Merged / Expanded HR Reference | VR-07 match on EmployeeID |
| Department From Reference | VR-04 (`IT`, blank, `Ops` → reference names) |
| Region From Reference | VR-10 |
| Approved* lists | Approved Values sheet (source of truth, not hard-coded) |
| Added Exception Reason / Validation Status | VR-01, 02, 03, 04, 05, 08, 10 checks; VR-05 → Excluded, other failures → Flagged |

Expected result: 11 rows: 10 `Valid`, plus E1008 `Excluded`, with reason `VR-05 TerminationDate before HireDate`.

## Workforce_Exceptions

```powerquery
let
    Source = Workforce_Validated,
    #"Kept Exception Records" = Table.SelectRows(Source, each [ValidationStatus] <> "Valid")
in
    #"Kept Exception Records"
```

## Workforce_Final (Step 4: the CSV and PivotTable source)

```powerquery
let
    Source = Workforce_Validated,
    #"Removed Excluded Records" = Table.SelectRows(Source, each [ValidationStatus] <> "Excluded"),
    #"Selected Output Fields" = Table.SelectColumns(#"Removed Excluded Records", {"EmployeeID", "EmployeeName", "Department", "Location", "Region", "EmploymentStatus", "EmploymentType", "HireDate", "TerminationDate", "FTE", "ValidationStatus"}),
    #"Formatted Dates YYYY-MM-DD" = Table.TransformColumns(#"Selected Output Fields", {{"HireDate", each Date.ToText(_, "yyyy-MM-dd"), type text}, {"TerminationDate", each if _ = null then null else Date.ToText(_, "yyyy-MM-dd"), type text}})
in
    #"Formatted Dates YYYY-MM-DD"
```

Dates are converted to text here so the CSV keeps `YYYY-MM-DD` (VR-09) whatever your Windows date locale is.

---

## Exporting the CSV

1. On the `Final` sheet, select the FTE column and set **Format Cells > Number, 1 decimal place**, so it exports `1.0` / `0.5`.
2. Right-click the `Final` sheet tab > **Move or Copy… > (new book)** and tick **Create a copy**.
3. In the new book: **File > Save As > CSV UTF-8 (Comma delimited)**, file name `Final_Analysis_Ready_Workforce.csv`.
4. Don't reopen and re-save the CSV in Excel, because Excel would reformat the dates.
5. Check: open the CSV in Notepad. It should match the reference `Final_Analysis_Ready_Workforce.csv` line for line: 1 header + 10 rows.

## PivotTable (Step 4)

1. Click inside the `Workforce_Final` table > **Insert > PivotTable > From Table/Range**. Table should be `Workforce_Final`. Place it on a new sheet `Pivot`, and tick **Add this data to the Data Model** (this enables Distinct Count).
2. **Rows:** Department
3. **Values:** EmployeeID → Value Field Settings > **Distinct Count**, renamed `Headcount`. Then FTE → **Sum**, renamed `Total FTE`, number format 0.0.
4. **Filters:** EmploymentStatus (leave as **All**).
5. Before capturing, check the data source: **PivotTable Analyze > Change Data Source** must show `Workforce_Final`.

Expected values:

| Department | Headcount | Total FTE |
|---|---|---|
| Customer Support | 1 | 1.0 |
| Finance | 2 | 2.0 |
| Human Resources | 1 | 1.0 |
| Information Technology | 1 | 1.0 |
| Marketing | 2 | 2.0 |
| Operations | 1 | 1.0 |
| Sales | 2 | 1.5 |
| **Grand Total** | **10** | **9.5** |

Headcount and FTE follow the **Metric Reference** sheet: distinct EmployeeID and sum of FTE over the final dataset. The total includes E1003 (Terminated 2026-07-18), because no supplied rule removes terminated employees from the reporting population. Set the EmploymentStatus filter to `Active` and the PivotTable shows 9 / 8.5. You can show that in a second capture, but the rule-defined figures are 10 / 9.5.

## Screens to capture for the PDF

1. `Recon_IDs` and `Recon_Field_Compare` result tables (Reconciliation)
2. Power Query editor with the `Workforce_Validated` query selected: Queries pane + Applied Steps pane visible (Power Query Transformations)
3. `Workforce_Validated` result showing ValidationStatus / ExceptionReason columns (Cleansing and Validation)
4. `Workforce_Exceptions` sheet showing E1008 (Exceptions)
5. `Workforce_Final` Applied Steps + result table (Power Query Transformations / final output)
6. PivotTable with the Fields pane visible (PivotTable Summary)
