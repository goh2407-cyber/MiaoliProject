---
name: manage_record_fields
description: Guide for adding or modifying fields in the Domestic Violence Records system. Covers Backend (GAS) and Frontend (HTML/JS) changes, and Manual Deployment.
---

# Manage Record Fields (SOP)

This skill provides a strict checklist for modifying the data schema of the **Domestic Violence Records** system.
Because this project uses Google Apps Script (GAS) with a tightly coupled spreadsheet backend, missing any step will result in data corruption or write failures.

## 📋 1. Backend Modifications (`RecordService.gs`)

Open `RecordService.gs` and modify the following areas in order:

### A. Update Field Definitions
Find the `RECORD_FIELDS` constant (approx. line 7).
- **Add** the new field name to the array.
- **CRITICAL**: The order MUST match your Google Sheet columns exactly.
- *Tip*: If adding a new column, usually append it before the system fields (`建立時間`...).

```javascript
const RECORD_FIELDS = [
  // ... existing fields ...
  'NewFieldName', // [NEW] Add this
  '建立時間', '修改時間', ...
];
```

### B. Update `createRecord` Function
Find `function createRecord(recordData)` (approx. line 24).
- Locate the `rowData` array definition.
- Add the mapping for your new field at the **same index position** as in `RECORD_FIELDS`.

```javascript
const rowData = [
  // ...
  recordData['NewFieldName'] || '', // [NEW] Add this
  // ...
];
```

### C. Update `updateRecord` Function
Find `function updateRecord(recordId, updateData)` (approx. line 94).
- Locate the `updatedRow` array definition.
- Add the update logic at the **same index position**.

```javascript
const updatedRow = [
  // ...
  updateData['NewFieldName'] !== undefined ? updateData['NewFieldName'] : existingRow[INDEX], // [NEW] replacing INDEX with the correct number
  // ...
];
```
> **Note**: You must count the index manually or ensure it aligns with `RECORD_FIELDS`.

### D. Update Validation (Optional)
If the field is required, update `validateRecordData` (approx. line 439) and add it to `requiredFields`.

---

## 🖥️ 2. Frontend Modifications

### A. HTML Update (`index.html`)

1.  **Add to Modal Form (`#recordForm`)**:
    - Add a new `<div class="form-group">` with a `label` and `input/select`.
    - **IMPORTANT**: The `id` of the input **MUST** match the `RECORD_FIELDS` name exactly.

    ```html
    <div class="form-group">
      <label>New Field Name</label>
      <input type="text" id="NewFieldName">
    </div>
    ```

2.  **Add to Preview Panel (`#previewContent`)**:
    - Add a display block in the appropriate `.preview-section`.
    - Remember to add `onchange="markPreviewChanged()"` to the input.

    ```html
    <div class="preview-item">
      <span>New Field Name</span>
      <input type="text" class="preview-edit-field" id="previewNewFieldName" onchange="markPreviewChanged()">
    </div>
    ```

### B. JavaScript Update (`scripts.html`)

1.  **For Select/Dropdown Fields**:
    - Find `function initializeFormOptions()` (approx. line 196).
    - Add your field name to the `selectFields` array. This ensures it gets populated with options from the Google Sheet.

    ```javascript
    const selectFields = [
        // ... existing fields ...
        'NewFieldName' // [NEW]
    ];
    ```

2.  **For Hierarchical Fields**:
    - You must define specific logic in `initSingleHierarchicalFields` or `initHierarchicalFields`. Refer to existing implementations like `服務主題`.

---

## 🚀 3. Manual Deployment (No Clasp)

Since `clasp` is not available, you must manually update the code in the Google Apps Script editor.

1.  **Copy Code**:
    - Copy the full content of the modified file (e.g., `RecordService.gs`).
2.  **Paste to GAS**:
    - Go to your Google Sheet -> Extensions -> Apps Script.
    - Open the corresponding file (`RecordService.gs`).
    - Select All (Ctrl+A) -> Delete -> Paste (Ctrl+V).
    - **Save** (Ctrl+S).
3.  **Repeat** for all modified files (`index.html`, `scripts.html`).
4.  **Reload Web App**:
    - Perform a new **Deployment** if you want to push to the stable URL, or use the **Test Deployment** URL to verify changes immediately.

## ⚠️ 4. Google Sheet Update

1.  **Open the Spreadsheet**.
2.  **Locate Column**: Go to the row with headers (usually Row 1).
3.  **Insert Column**: Add the new column header `NewFieldName`.
4.  **Order Check**: Ensure the column position matches exactly with `RECORD_FIELDS` in `RecordService.gs`.
