
/**
* @OnlyCurrentDoc
* Manages an inventory system by linking a Google Sheet to Google Forms.
* Supports multiple sheets/forms (CC Connectors, CC Supplies).
* Contains functions to:
* 1. Dynamically update dropdown questions in forms based on sheet inventory.
* 2. Process form submissions to automatically adjust inventory quantities in the sheet.
*/


// --- MASTER CONFIGURATION ---
// Per-form config. responseSheetName = sheet where this form's responses are stored (used to detect which form was submitted).
const CONFIGS = {
    connectors: {
      sheetName:           "CC Connectors",
      formId:              "1v7knzJcb-NMW2d3agCk-97Sx19IOTNS-m-PKvs0haYg",
      helperColumn:        "P",
      dropdownQuestionTitle: "Select the Bin",
      actionQuestionTitle: "Are you taking or adding items?",
      quantityQuestionTitle: "Quantity",
      inventoryBinColumn:  1,
      inventoryQuantityColumn: 10,
      headersRow:          3,
      responseSheetName:   "CONNECTORS - IN / OUT"   // Sheet that receives this form's responses
    },
    supplies: {
      sheetName:           "CC Supplies",
      // Use the form ID from the form's EDIT URL (open form → ⋮ → Edit form → URL is .../d/FORM_ID/edit). The view URL (/d/e/.../viewform) may not work.
      formId:              "1yqTiKLfXiKFdbpbkpd0kVYV_1YfpykDX03vNwjHuGwU",
      helperColumn:        "P",
      dropdownQuestionTitle: "Select the Bin",
      actionQuestionTitle: "Are you taking or adding items?",
      quantityQuestionTitle: "Quantity",
      inventoryBinColumn:  1,
      inventoryQuantityColumn: 10,
      headersRow:          3,
      responseSheetName:   "SUPPLIES - IN / OUT"   // Set to the actual response sheet name for CC Supplies form
    }
  };
  // --------------------------
  
  
  /**
  * Updates the dropdown list in a Google Form with bin values from the given sheet config.
  * @param {Object} config - Config object (e.g. CONFIGS.connectors or CONFIGS.supplies)
  */
  function updateFormDropdownForConfig(config) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) throw new Error(`Sheet "${config.sheetName}" not found.`);
  
    const range = sheet.getRange(`${config.helperColumn}4:${config.helperColumn}`);
    const values = range.getValues().flat().filter(String);
  
    if (values.length === 0) {
      Logger.log(`No values found in ${config.sheetName} helper column to update the form.`);
      return;
    }
  
    let form;
    try {
      form = FormApp.openById(config.formId);
    } catch (openErr) {
      throw new Error(
        `Could not open form (ID: ${config.formId}). ` +
        "Use the form's EDIT URL (.../d/FORM_ID/edit) and ensure this account has Editor access. " +
        openErr.message
      );
    }
  
    const items = form.getItems(FormApp.ItemType.LIST);
    const dropdownItem = items.find(item => item.getTitle() === config.dropdownQuestionTitle);
  
    if (dropdownItem) {
      dropdownItem.asListItem().setChoiceValues(values);
      Logger.log(`Dropdown "${config.dropdownQuestionTitle}" updated successfully for ${config.sheetName}.`);
    } else {
      const titles = items.map(i => i.getTitle()).join(", ");
      throw new Error(
        `Question "${config.dropdownQuestionTitle}" not found. List question titles in this form: ${titles || "(none)"}`
      );
    }
  }
  
  /**
  * Updates the CC Connectors form dropdown with bin numbers/names from the CC Connectors sheet.
  */
  function updateFormDropdown() {
    try {
      updateFormDropdownForConfig(CONFIGS.connectors);
    } catch (e) {
      Logger.log(`Error updating Connectors form dropdown: ${e.message}`);
    }
  }
  
  /**
  * Updates the CC Supplies form dropdown with bin numbers/names from the CC Supplies sheet.
  */
  function updateFormDropdownSupplies() {
    try {
      updateFormDropdownForConfig(CONFIGS.supplies);
    } catch (e) {
      Logger.log(`Error updating Supplies form dropdown: ${e.message}`);
    }
  }
  
  /**
  * Updates both the CC Connectors and CC Supplies form dropdowns.
  */
  function updateAllFormDropdowns() {
    try {
      updateFormDropdownForConfig(CONFIGS.connectors);
    } catch (e) {
      Logger.log(`Error updating Connectors form dropdown: ${e.message}`);
    }
    try {
      updateFormDropdownForConfig(CONFIGS.supplies);
    } catch (e) {
      Logger.log(`Error updating Supplies form dropdown: ${e.message}`);
    }
  }
  
  /**
  * Runs when the spreadsheet is edited. If the edit was on the CC Supplies sheet,
  * updates the CC Supplies form dropdown. Add a trigger: Triggers → Add Trigger →
  * function: onEdit, event: From spreadsheet → On edit.
  * @param {Object} e The event object passed by the "On edit" trigger.
  */
  function onEdit(e) {
    if (!e || !e.range) return;
    const sheetName = e.range.getSheet().getName();
    if (sheetName === CONFIGS.supplies.sheetName) {
      updateFormDropdownSupplies();
    }
  }
  
  /**
  * Returns the config for the form that submitted (based on which sheet received the response).
  * @param {Object} e The event object passed by the "On form submit" trigger.
  * @returns {Object|null} Config object or null if no match.
  */
  function getConfigForSubmitEvent(e) {
    const responseSheetName = e.range.getSheet().getName();
    for (const key in CONFIGS) {
      if (CONFIGS[key].responseSheetName === responseSheetName) {
        return CONFIGS[key];
      }
    }
    return null;
  }
  
  /**
  * Runs when a form is submitted, finds the corresponding item, and updates the quantity.
  * Works for both CC Connectors and CC Supplies based on which sheet received the response.
  * @param {Object} e The event object passed by the "On form submit" trigger.
  */
  function onFormSubmit(e) {
    try {
      const config = getConfigForSubmitEvent(e);
      if (!config) {
        Logger.log(`No config found for response sheet "${e.range.getSheet().getName()}". Check CONFIGS.responseSheetName.`);
        return;
      }
  
      const responses = e.namedValues;
      const itemLabel = responses[config.dropdownQuestionTitle][0];
      const action = responses[config.actionQuestionTitle][0];
      const quantity = parseInt(responses[config.quantityQuestionTitle][0], 10);
  
      if (isNaN(quantity)) {
        throw new Error("Submitted quantity is not a valid number.");
      }
  
      const binNumber = itemLabel.split(' - ')[0].trim();
  
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const inventorySheet = ss.getSheetByName(config.sheetName);
      if (!inventorySheet) throw new Error(`Sheet "${config.sheetName}" not found.`);
  
      const dataRange = inventorySheet.getRange(
        config.headersRow + 1, 1,
        inventorySheet.getLastRow(),
        inventorySheet.getLastColumn()
      );
      const inventoryData = dataRange.getValues();
  
      let targetRow = -1;
      for (let i = 0; i < inventoryData.length; i++) {
        if (inventoryData[i][config.inventoryBinColumn - 1].toString().trim() === binNumber) {
          targetRow = i;
          break;
        }
      }
  
      if (targetRow !== -1) {
        const sheetRowIndex = targetRow + config.headersRow + 1;
        const quantityCell = inventorySheet.getRange(sheetRowIndex, config.inventoryQuantityColumn);
        const currentQuantity = quantityCell.getValue();
  
        let newQuantity;
        if (action === "Taking from inventory") {
          newQuantity = currentQuantity - quantity;
        } else {
          newQuantity = currentQuantity + quantity;
        }
  
        quantityCell.setValue(newQuantity);
        Logger.log(`Updated ${config.sheetName} Bin ${binNumber}: Quantity changed from ${currentQuantity} to ${newQuantity}.`);
      } else {
        throw new Error(`Could not find Bin Number "${binNumber}" in the inventory sheet "${config.sheetName}".`);
      }
    } catch (e) {
      Logger.log(`Error processing form submission: ${e.message}`);
    }
  }
  
  
  /**
  * Debug: Opens the CC Supplies form and logs all question titles.
  * Run this to verify the form ID works and to see the exact dropdown question title (use it in dropdownQuestionTitle if different).
  */
  function debugListSuppliesFormQuestions() {
    const config = CONFIGS.supplies;
    try {
      const form = FormApp.openById(config.formId);
      const items = form.getItems();
      Logger.log("CC Supplies form opened. Question titles:");
      items.forEach(function(item) {
        Logger.log("  - " + item.getTitle() + " (type: " + item.getType() + ")");
      });
    } catch (e) {
      Logger.log("Error: " + e.message);
      Logger.log("Get the form ID from the form's EDIT URL: open the form, ⋮ → Edit form, copy the ID from .../d/FORM_ID/edit. Ensure this account has Editor access.");
    }
  }
  
  /**
  * A temporary function to debug form submissions.
  * It logs the exact question titles and answers received from the form.
  */
  function debugOnSubmit(e) {
   // Log the raw data object from the form submission
   Logger.log("--- Form Data Received ---");
   Logger.log(JSON.stringify(e.namedValues, null, 2));
    // Log just the question titles (the keys)
   Logger.log("--- Question Titles ---");
   Logger.log(Object.keys(e.namedValues));
  }
  
  
  