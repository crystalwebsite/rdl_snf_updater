# Power BI RDL Snowflake Connection Updater

This Python utility automates Snowflake server/warehouse changes in Power BI `.rdl` files.

It is designed for the Power BI RDL `CommandText` structure where the command is a JSON payload containing a `MashupDocument`, for example:

```text
Source = Snowflake.Databases(\"myprod.snowflakecomputing.com\", \"MY_WAREHOUSE\", [Implementation = \"2.0\"])
```

and connection metadata such as:

```json
"ConnectionOverrides":[
  {
    "Path":"myprod.snowflakecomputing.com;MY_WAREHOUSE",
    "Kind":"Snowflake"
  }
]
```

## What it does

For every `.rdl` file placed in the input folder, the program:

1. Reads and validates the RDL as XML.
2. Displays all `DataSet/@Name` values under `<DataSets>`.
3. Scans every dataset's `<CommandText>`.
4. Finds `Source = Snowflake.Databases(...)` even when the M-code quotes are JSON-escaped as `\"...\"`.
5. Replaces the first positional parameter with the configured Snowflake server.
6. Replaces the second positional parameter with the configured Snowflake warehouse.
7. Synchronizes a matching `ConnectionOverrides[].Path` from:

   ```text
   old_server;OLD_WAREHOUSE
   ```

   to:

   ```text
   new_server;NEW_WAREHOUSE
   ```

8. Validates the transformed RDL again.
9. Writes the result to the output folder.
10. Appends detailed before/after audit rows to an Excel workbook.
11. Deletes the original input RDL **only after both the output RDL and Excel audit log are successfully written**.

If no `Snowflake.Databases(...)` source exists in the RDL, the original file is kept in the input folder.

## Example transformation

Before:

```text
Source = Snowflake.Databases(\"myprod.snowflakecomputing.com\", \"MY_WAREHOUSE\", [Implementation = \"2.0\"])
```

After:

```text
Source = Snowflake.Databases(\"mytest.snowflakecomputing.com\", \"TEST_WH\", [Implementation = \"2.0\"])
```

The corresponding connection override is also synchronized.

Before:

```json
"Path":"myprod.snowflakecomputing.com;MY_WAREHOUSE"
```

After:

```json
"Path":"mytest.snowflakecomputing.com;TEST_WH"
```

## Folder structure

```text
rdl_snowflake_updater/
├─ rdl_snowflake_updater.py
├─ config.json
├─ requirements.txt
├─ README.md
├─ input/
├─ output/
└─ logs/
```

The Python program creates `input`, `output`, and `logs` automatically when needed.

## Setup on Windows

From PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Configuration

Edit `config.json`:

```json
{
  "input_dir": "input",
  "output_dir": "output",
  "audit_file": "logs/rdl_modification_log.xlsx",
  "snowflake_server": "mytest.snowflakecomputing.com",
  "snowflake_warehouse": "TEST_WH",
  "overwrite": false
}
```

## Run

Place one or more `.rdl` files in `input`, then run:

```powershell
python rdl_snowflake_updater.py
```

You can also override values from the command line:

```powershell
python rdl_snowflake_updater.py `
  --server "mytest.snowflakecomputing.com" `
  --warehouse "TEST_WH"
```

To permit replacement of an existing output RDL:

```powershell
python rdl_snowflake_updater.py `
  --server "mytest.snowflakecomputing.com" `
  --warehouse "TEST_WH" `
  --overwrite
```

## Console example

```text
File: MyReport.rdl
DataSets found:
  1. DefaultPeriod
  2. SalesData
  3. CustomerData

Snowflake.Databases source(s) found: 3
Source value change(s): 3
ConnectionOverrides.Path match(es): 3
ConnectionOverrides.Path change(s): 3
```

## Excel audit log

The workbook is stored by default at:

```text
logs/rdl_modification_log.xlsx
```

The `ModificationLog` worksheet records:

- Timestamp
- File Name
- Dataset Name
- Modification Type
- Status
- Before Server
- After Server
- Before Warehouse
- After Warehouse
- Before Value
- After Value
- Before Code Snapshot
- After Code Snapshot
- Input Path
- Output Path
- Original SHA256
- Modified SHA256

A Snowflake dataset can therefore produce two audit rows:

1. `Snowflake.Databases Source`
2. `ConnectionOverrides.Path`

This gives a direct before/after record of both the M-code connection parameters and the Power BI connection metadata.

## Safety behavior

The original file is deleted only in this order:

```text
Read input RDL
     │
     ▼
Validate XML
     │
     ▼
Find Snowflake source
     │
     ▼
Transform values
     │
     ▼
Validate transformed XML
     │
     ▼
Write output RDL successfully
     │
     ▼
Write Excel audit successfully
     │
     ▼
Delete original input RDL
```

If output or audit writing fails, the original file remains available.

## Important matching rule

`ConnectionOverrides.Path` is changed only when its original `server;warehouse` pair matches a `Snowflake.Databases(server, warehouse, ...)` pair found in the same `<CommandText>`. This prevents unrelated connection paths from being modified accidentally.
