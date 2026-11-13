// This file runs inside a WPS add-in (not Node), with `Application` available.
return {
  address: Application.Selection.Address,
  workbook: Application.ActiveWorkbook?.Name ?? null
};
