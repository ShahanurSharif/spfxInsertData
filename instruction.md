# Insert Data Web Part Instructions

## 1. How to Install

**Prerequisites:**
- You must be a SharePoint administrator or have permissions to upload apps to the App Catalog.

**Steps:**
1. Upload the `spfx-insert-data.sppkg` file (found in `sharepoint/solution/`) to your SharePoint App Catalog.
2. Approve and deploy the solution if prompted.
3. Go to your target SharePoint site, open **Site Contents**, and add the app if it’s not already present.
4. Edit a page and add the “Insert Data” web part.

---

## 2. How to Use

- Click the **Create Item** button to open the form.
- Fill in the **Title**, **Description**, and select a **Letter** from the dropdown.
- Click **Submit** to add a new FAQ item.
- Existing items are shown in a table. Use **Edit** to update or **Delete** to remove items.
- Success and error messages will appear at the top of the web part.

---

## 3. How to Create the FAQ List and Fields

**You must create a SharePoint list named `FAQ` with the following fields:**

1. **Title**  
   - Type: Single line of text  
   - (This is the default field in every SharePoint list.)

2. **Description**  
   - Internal Name: `Body` (case-sensitive, capital "B")  
   - Display Name: Description  
   - Type: Multiple lines of text (recommended) or Single line of text  
   - To ensure the internal name is `Body`, create a new column and name it `Description` when creating the list. SharePoint will use `Body` as the internal name for the first multiline text field in a custom list.

3. **Letter**  
   - Type: Choice  
   - Choices: A, B, C (or any set you want)  
   - Internal Name: Letter

**How to create the list and fields:**
1. Go to your SharePoint site.
2. Click **New > List** and name it `FAQ`.
3. In the FAQ list, click **Add column > Multiple lines of text**.  
   - Name it `Description`.  
   - (After creation, verify the internal name is `Body` by clicking the column, then checking the URL for `Field=Body`.)
4. Click **Add column > Choice**.  
   - Name it `Letter`.  
   - Add choices (A, B, C, etc.).

**Note:**
- The web part expects the internal name for Description to be `Body`. If you use a different name, update the code accordingly.
- The web part will not work if the field names do not match these internal names.

---

## 4. Troubleshooting

- If you see errors about missing fields, double-check the internal names in your list settings.
- If you update the list structure, you may need to refresh the page or re-add the web part.

---
