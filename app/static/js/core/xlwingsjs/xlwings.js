import { xlAlert } from "./alert.js";
import { getAccessToken } from "./auth.js";
export { getAccessToken };
import {
  getActiveBookName,
  printSupportedApiVersions,
  getCultureInfoName,
  getDateFormat,
  showGlobalError,
  showGlobalStatus,
  hideGlobalError,
  hideGlobalStatus,
} from "./utils.js";
export { getActiveBookName, getCultureInfoName, getDateFormat };
import { pyodideReadyPromise } from "./lite.js";
import { registerSheetButtons } from "./sheet-buttons.js";

// Prints the supported API versions into the Console
printSupportedApiVersions();

// Namespace
const xlwings = {
  runPython,
  getAccessToken,
  getActiveBookName,
  getBookData,
  runActions,
  pyodideReadyPromise,
  getCultureInfoName,
  getDateFormat,
  init,
  registerSheetButtons,
  showGlobalError,
  hideGlobalError,
  showGlobalStatus,
  hideGlobalStatus,
  registerCallback,
};
globalThis.xlwings = xlwings;

// Hook up buttons with the click event upon loading xlwings.js
document.addEventListener("DOMContentLoaded", init);

export async function init() {
  await xlwings.pyodideReadyPromise;
  // Handle unsupported browsers (IE/Edge Legacy)
  if (
    navigator.userAgent.indexOf("Trident") !== -1 ||
    navigator.userAgent.indexOf("Edge") !== -1
  ) {
    showGlobalError(
      "Error: This add-in will not run in your version of Office. Please upgrade " +
        "either to perpetual Office 2021 (or later) or to a Microsoft 365 account.",
    );
    return;
  }
  // Scripts meta
  let scriptsMeta = [];
  if (config.onLite && !config.isOfficialLiteAddin) {
    scriptsMeta = globalThis.liteCustomScriptsMeta();
  } else if (!config.isOfficialLiteAddin) {
    const metaUrl =
      window.location.origin +
      config.appPath +
      "/xlwings/custom-scripts-meta.json";
    try {
      const response = await axios.get(metaUrl);
      scriptsMeta = response.data;
    } catch (error) {
      console.error("Error fetching script metadata:", error);
    }
  }

  // xw-click registration
  const elements = document.querySelectorAll("[xw-click]");
  elements.forEach((element) => {
    // Prevent duplicate initialization when loading partials via htmx
    if (element.hasAttribute("xw-click-initialized")) return;
    element.setAttribute("xw-click-initialized", "true");
    element.addEventListener("click", async (event) => {
      // Clean up error messages
      const globalErrorAlert = document.querySelector("#global-error-alert");
      if (globalErrorAlert) {
        globalErrorAlert.classList.add("d-none");
      }
      element.setAttribute("disabled", "true");
      // Spinner
      const spinner = document.createElement("span");
      spinner.className = "spinner-border spinner-border-sm text-white";
      spinner.setAttribute("role", "status");
      spinner.setAttribute("aria-hidden", "true");
      element.appendChild(spinner);

      let token =
        typeof globalThis.getAuth === "function"
          ? await globalThis.getAuth()
          : "";
      let scriptName = element.getAttribute("xw-click");

      // Config
      let xwConfig = element.getAttribute("xw-config")
        ? JSON.parse(element.getAttribute("xw-config"))
        : {};
      // Find the script config that matches the current scriptName
      const matchingMeta = scriptsMeta.find(
        (meta) => meta.function_name === scriptName,
      );
      if (matchingMeta) {
        // Override xwConfig with the matched meta's config
        xwConfig = {
          ...xwConfig,
          exclude: matchingMeta.exclude || "",
          include: matchingMeta.include || "",
        };
      }
      // Call runPython and restore button default state
      await runPython({
        ...xwConfig,
        scriptName: scriptName,
        auth: token,
        errorDisplayMode: "taskpane",
      });
      element.removeChild(spinner);
      element.removeAttribute("disabled");
    });
  });
  // Handle sheet buttons
  await registerSheetButtons(scriptsMeta);
}

const version = config.xlwingsVersion;

globalThis.callbacks = {};
export async function runPython({
  scriptName = "",
  auth = "",
  include = "",
  exclude = "",
  headers = {},
  errorDisplayMode = "alert",
} = {}) {
  await Office.onReady();
  try {
    await Excel.run(async (context) => {
      // console.log(payload);
      let payload = await getBookData(
        {
          auth,
          include,
          exclude,
          headers,
        },
        context,
      );
      let rawData;
      if (config.onLite) {
        // xlwings Wasm
        await pyodideReadyPromise;
        rawData = await globalThis.liteCustomScriptsCall(payload, scriptName);
        if (rawData.error) {
          console.error(rawData.details);
          throw new Error(rawData.error);
        }
      } else {
        // xlwings Server
        let url =
          window.location.origin +
          config.appPath +
          `/xlwings/custom-scripts-call/${scriptName}`;
        try {
          const response = await axios.post(url, payload, {
            headers: headers,
            timeout: config.requestTimeout * 1000,
          });
          rawData = response.data;
        } catch (error) {
          // TODO: align error handling with xlwings Lite
          if (error.response) {
            throw (
              (error.response.data && error.response.data.detail) ||
              (error.response.data && error.response.data.error) ||
              (typeof error.response.data === "object"
                ? JSON.stringify(error.response.data)
                : error.response.data) ||
              error.response.statusText ||
              "Unknown server error"
            );
          } else if (error.request) {
            throw "No response received from server";
          } else {
            throw error.message;
          }
        }
      }
      // console.log(rawData);

      // Run Functions
      // Note that Pyodide returns undefined, so use != and == rather than !== and ===
      if (rawData != null) {
        await runActions(rawData, context);
      }
    });
  } catch (error) {
    console.error(error);
    if (errorDisplayMode === "alert") {
      await xlAlert(error, "Error", "ok", "critical", "");
    } else {
      const globalErrorAlert = document.querySelector("#global-error-alert");
      if (globalErrorAlert) {
        globalErrorAlert.classList.remove("d-none");
        globalErrorAlert.querySelector("span").textContent = error;
      }
    }
  }
}

// Helpers
async function getSelectedRangeAddress(context) {
  let selectionAddress = null;
  try {
    let selection = context.workbook.getSelectedRange().load("address");
    await context.sync();
    selectionAddress = selection.address.split("!").pop();
  } catch (error) {
    // No range is selected (e.g., a shape is selected)
  }
  return selectionAddress;
}

async function getBookData(
  { auth = "", include = "", exclude = "", headers = {} } = {},
  context = null,
) {
  // Context
  let bookData;
  if (!context) {
    await Excel.run(async (innerContext) => {
      bookData = await getBookData(
        {
          auth,
          include,
          exclude,
          headers,
        },
        innerContext,
      );
    });
    return bookData;
  }

  // workbook
  const workbook = context.workbook;
  workbook.load("name");

  // sheets
  let worksheets = workbook.worksheets;
  worksheets.load("items/name");
  await context.sync();
  let sheets = worksheets.items;

  // Config
  let configSheet = worksheets.getItemOrNullObject("xlwings.conf");
  await context.sync();
  let config = {};
  if (!configSheet.isNullObject) {
    const configRange = configSheet
      .getRange("A1")
      .getSurroundingRegion()
      .load("values");
    await context.sync();
    const configValues = configRange.values;
    configValues.forEach((el) => (config[el[0].toString()] = el[1].toString()));
  }

  if (auth === "") {
    auth = config["AUTH"] || "";
  }

  if (include === "") {
    include = config["INCLUDE"] || "";
  }
  let includeArray = [];
  if (include !== "") {
    includeArray = include.split(",").map((item) => item.trim());
  }

  if (exclude === "") {
    exclude = config["EXCLUDE"] || "";
  }
  let excludeArray = [];
  if (exclude !== "") {
    excludeArray = exclude.split(",").map((item) => item.trim());
  }
  if (includeArray.length > 0 && excludeArray.length > 0) {
    throw "Either use 'include' or 'exclude', but not both!";
  }
  if (includeArray.length > 0) {
    sheets.forEach((sheet) => {
      if (!includeArray.includes(sheet.name)) {
        excludeArray.push(sheet.name);
      }
    });
  }

  if (Object.keys(headers).length === 0) {
    for (const property in config) {
      if (property.toLowerCase().startsWith("header_")) {
        headers[property.substring(7)] = config[property];
      }
    }
  }
  if (!("Authorization" in headers) && auth.length > 0) {
    headers["Authorization"] = auth;
  }

  // Standard headers
  headers["Content-Type"] = "application/json";

  // Request payload
  let payload = {};
  payload["client"] = "Office.js";
  payload["version"] = version;
  let activeSheet = worksheets.getActiveWorksheet().load("position");
  await context.sync();

  // Cell selection address
  const selectionAddress = await getSelectedRangeAddress(context);

  payload["book"] = {
    name: workbook.name,
    active_sheet_index: activeSheet.position,
    selection: selectionAddress,
  };

  // Names (book scope)
  let names = [];
  const namedItems = context.workbook.names.load("name, type");
  await context.sync();

  for (const namedItem of namedItems.items) {
    // Currently filtering to named ranges
    if (namedItem.type === "Range") {
      // Names pointing to multiple Ranges return null
      let range = namedItem.getRangeOrNullObject();
      await context.sync();
      names.push({
        name: namedItem.name,
        sheet: range.isNullObject ? null : range.worksheet.load("position"),
        range: range.isNullObject ? null : range.load("address"),
        scope_sheet_name: null,
        scope_sheet_index: null,
        book_scope: true, // workbook.names contains only workbook scope!
      });
    }
  }

  await context.sync();

  let names2 = [];
  names.forEach((namedItem, ix) => {
    names2.push({
      name: namedItem.name,
      sheet_index: namedItem.sheet ? namedItem.sheet.position : null,
      address: namedItem.range
        ? namedItem.range.address.split("!").pop()
        : null,
      scope_sheet_name: null,
      scope_sheet_index: null,
      book_scope: namedItem.book_scope,
    });
  });

  payload["names"] = names2;

  // Sheets
  payload["sheets"] = [];
  let sheetsLoader = [];
  sheets.forEach((sheet) => {
    sheet.load("name names");
    let lastCell;
    if (excludeArray.includes(sheet.name)) {
      lastCell = null;
    } else if (sheet.getUsedRange() !== undefined) {
      lastCell = sheet.getUsedRange().getLastCell().load("address");
    } else {
      lastCell = sheet.getRange("A1").load("address");
    }
    sheetsLoader.push({
      sheet: sheet,
      lastCell: lastCell,
    });
  });

  await context.sync();

  sheetsLoader.forEach((item, ix) => {
    if (!excludeArray.includes(item["sheet"].name)) {
      let range;
      range = item["sheet"]
        .getRange(`A1:${item["lastCell"].address}`)
        .load("values, numberFormatCategories");
      sheetsLoader[ix]["range"] = range;
      // Names (sheet scope)
      sheetsLoader[ix]["names"] = item["sheet"].names.load("name, type");
    }
  });

  await context.sync();

  // Names (sheet scope)
  let namesSheetScope = [];
  for (const item of sheetsLoader) {
    if (!excludeArray.includes(item["sheet"].name)) {
      for (const namedItem of item["names"].items) {
        // Currently filtering to named ranges
        if (namedItem.type === "Range") {
          let range = namedItem.getRangeOrNullObject();
          await context.sync();
          namesSheetScope.push({
            name: namedItem.name,
            sheet: range.isNullObject ? null : range.worksheet.load("position"),
            range: range.isNullObject ? null : range.load("address"),
            scope_sheet: namedItem.worksheet.load("name, position"),
            book_scope: false,
          });
        }
      }
    }
  }

  await context.sync();

  let namesSheetsScope2 = [];
  for (const namedItem of namesSheetScope) {
    namesSheetsScope2.push({
      name: namedItem.name,
      sheet_index: namedItem.sheet ? namedItem.sheet.position : null,
      address: namedItem.range
        ? namedItem.range.address.split("!").pop()
        : null,
      scope_sheet_name: namedItem.scope_sheet.name,
      scope_sheet_index: namedItem.scope_sheet.position,
      book_scope: namedItem.book_scope,
    });
  }

  // Add sheet scoped names to book scoped names
  payload["names"] = payload["names"].concat(namesSheetsScope2);

  // values
  for (let item of sheetsLoader) {
    let sheet = item["sheet"]; // TODO: replace item["sheet"] with sheet
    let values;
    if (excludeArray.includes(item["sheet"].name)) {
      values = [[]];
    } else {
      values = item["range"].values;
      if (Office.context.requirements.isSetSupported("ExcelApi", "1.12")) {
        // numberFormatCategories requires Excel 2021/365
        // i.e., dates aren't transformed to Python's datetime in Excel <=2019
        let categories = item["range"].numberFormatCategories;
        // Handle dates
        // https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/excel-samples#dates
        values.forEach((valueRow, rowIndex) => {
          const categoryRow = categories[rowIndex];
          valueRow.forEach((value, colIndex) => {
            const category = categoryRow[colIndex];
            if (
              (category.toString() === "Date" ||
                category.toString() === "Time") &&
              typeof value === "number"
            ) {
              values[rowIndex][colIndex] = new Date(
                Math.round((value - 25569) * 86400 * 1000),
              ).toISOString();
            }
          });
        });
      }
    }
    // Tables
    let tablesArray = [];
    if (!excludeArray.includes(item["sheet"].name)) {
      const tables = sheet.tables.load([
        "name",
        "showHeaders",
        "dataBodyRange",
        "showTotals",
        "style",
        "showFilterButton",
      ]);
      await context.sync();
      let tablesLoader = [];
      for (let table of sheet.tables.items) {
        tablesLoader.push({
          name: table.name,
          showHeaders: table.showHeaders,
          showTotals: table.showTotals,
          style: table.style,
          showFilterButton: table.showFilterButton,
          range: table.getRange().load("address"),
          dataBodyRange: table.getDataBodyRange().load("address"),
          headerRowRange: table.showHeaders
            ? table.getHeaderRowRange().load("address")
            : null,
          totalRowRange: table.showTotals
            ? table.getTotalRowRange().load("address")
            : null,
        });
      }
      await context.sync();
      for (let table of tablesLoader) {
        tablesArray.push({
          name: table.name,
          range_address: table.range.address.split("!").pop(),
          header_row_range_address: table.showHeaders
            ? table.headerRowRange.address.split("!").pop()
            : null,
          data_body_range_address: table.dataBodyRange.address.split("!").pop(),
          total_row_range_address: table.showTotals
            ? table.totalRowRange.address.split("!").pop()
            : null,
          show_headers: table.showHeaders,
          show_totals: table.showTotals,
          table_style: table.style,
          show_autofilter: table.showFilterButton,
        });
      }
    }

    // Pictures
    let picturesArray = [];
    if (!excludeArray.includes(item["sheet"].name)) {
      const shapes = sheet.shapes.load(["name", "width", "height", "type"]);
      await context.sync();
      for (let shape of sheet.shapes.items) {
        if (shape.type == Excel.ShapeType.image) {
          picturesArray.push({
            name: shape.name,
            height: shape.height,
            width: shape.width,
          });
        }
      }
    }

    payload["sheets"].push({
      name: item["sheet"].name,
      values: values,
      pictures: picturesArray,
      tables: tablesArray,
    });
  }
  return payload;
}

async function runActions(rawData, context = null) {
  if (typeof rawData === "string") {
    rawData = JSON.parse(rawData);
  }

  if (!context) {
    return await Excel.run(async (innerContext) => {
      await runActions(rawData, innerContext);
    });
  }

  const forceSync = ["sheet"];
  for (let action of rawData["actions"]) {
    await globalThis.callbacks[action.func](context, action);
    if (forceSync.some((el) => action.func.toLowerCase().includes(el))) {
      await context.sync();
    }
  }
}

async function getRange(context, action) {
  let sheets = context.workbook.worksheets.load("items");
  await context.sync();
  return sheets.items[action["sheet_position"]].getRangeByIndexes(
    action.start_row,
    action.start_column,
    action.row_count,
    action.column_count,
  );
}

async function getSheet(context, action) {
  let sheets = context.workbook.worksheets.load("items");
  await context.sync();
  return sheets.items[action.sheet_position];
}

async function getTable(context, action) {
  // Requires action.args[0] to be the table index
  let sheets = context.workbook.worksheets.load("items");
  const tables = sheets.items[action.sheet_position].tables.load("items");
  await context.sync();
  return tables.items[parseInt(action.args[0].toString())];
}

async function getShapeByType(context, sheetPosition, shapeIndex, shapeType) {
  let sheets = context.workbook.worksheets.load("items");
  const shapes = sheets.items[sheetPosition].shapes.load("items");
  await context.sync();
  const myshapes = shapes.items.filter((shape) => shape.type === shapeType);
  return myshapes[shapeIndex];
}

export function registerCallback(callback) {
  globalThis.callbacks[callback.name] = callback;
}

// Functions map
// Didn't find a way to use registerCallback so that webpack won't strip out these
// functions when optimizing
let funcs = {
  setValues: setValues,
  addSheet: addSheet,
  setSheetName: setSheetName,
  setAutofit: setAutofit,
  setRangeColor: setRangeColor,
  activateSheet: activateSheet,
  addHyperlink: addHyperlink,
  setNumberFormat: setNumberFormat,
  setPictureName: setPictureName,
  setPictureWidth: setPictureWidth,
  setPictureHeight: setPictureHeight,
  deletePicture: deletePicture,
  addPicture: addPicture,
  updatePicture: updatePicture,
  alert: alert,
  setRangeName: setRangeName,
  namesAdd: namesAdd,
  nameDelete: nameDelete,
  runMacro: runMacro,
  rangeDelete: rangeDelete,
  rangeInsert: rangeInsert,
  rangeSelect: rangeSelect,
  rangeClearContents: rangeClearContents,
  rangeClearFormats: rangeClearFormats,
  rangeGroup: rangeGroup,
  rangeUngroup: rangeUngroup,
  rangeClear: rangeClear,
  rangeAdjustIndent: rangeAdjustIndent,
  addTable: addTable,
  setTableName: setTableName,
  resizeTable: resizeTable,
  showAutofilterTable: showAutofilterTable,
  showHeadersTable: showHeadersTable,
  showTotalsTable: showTotalsTable,
  setTableStyle: setTableStyle,
  copyRange: copyRange,
  copyFromRange: copyFromRange,
  sheetDelete: sheetDelete,
  sheetClear: sheetClear,
  sheetClearFormats: sheetClearFormats,
  sheetClearContents: sheetClearContents,
  freezePaneAtRange: freezePaneAtRange,
  freezePaneUnfreeze: freezePaneUnfreeze,
  setFontProperty: setFontProperty,
};

Object.assign(globalThis.callbacks, funcs);

// Callbacks
async function setFontProperty(context, action) {
  let range = await getRange(context, action);
  let property = action.args[0];
  let value = action.args[1];
  if (property === "bold" || property === "italic") value = Boolean(value);
  range.format.font[property] = value;
  await context.sync();
}

async function setValues(context, action) {
  let range = await getRange(context, action);
  range.values = action.values;
  await context.sync();
}

async function rangeClearContents(context, action) {
  let range = await getRange(context, action);
  range.clear(Excel.ClearApplyTo.contents);
  await context.sync();
}

async function rangeClearFormats(context, action) {
  let range = await getRange(context, action);
  range.clear(Excel.ClearApplyTo.formats);
  await context.sync();
}

async function rangeClear(context, action) {
  let range = await getRange(context, action);
  range.clear(Excel.ClearApplyTo.all);
  await context.sync();
}

async function addSheet(context, action) {
  let sheet;
  if (action.args[1] != null) {
    sheet = context.workbook.worksheets.add(action.args[1].toString());
  } else {
    sheet = context.workbook.worksheets.add();
  }
  sheet.position = parseInt(action.args[0].toString());
}

async function setSheetName(context, action) {
  let sheets = context.workbook.worksheets.load("items");
  sheets.items[action.sheet_position].name = action.args[0].toString();
}

async function setAutofit(context, action) {
  if (action.args[0] === "columns") {
    let range = await getRange(context, action);
    range.format.autofitColumns();
  } else {
    let range = await getRange(context, action);
    range.format.autofitRows();
  }
}

async function setRangeColor(context, action) {
  let range = await getRange(context, action);
  range.format.fill.color = action.args[0].toString();
  await context.sync();
}

async function activateSheet(context, action) {
  let worksheets = context.workbook.worksheets;
  worksheets.load("items");
  await context.sync();
  worksheets.items[parseInt(action.args[0].toString())].activate();
}

async function addHyperlink(context, action) {
  let range = await getRange(context, action);
  let hyperlink = {
    textToDisplay: action.args[1].toString(),
    screenTip: action.args[2].toString(),
    address: action.args[0].toString(),
  };
  range.hyperlink = hyperlink;
  await context.sync();
}

async function setNumberFormat(context, action) {
  let range = await getRange(context, action);
  range.numberFormat = [[action.args[0].toString()]];
}

async function setPictureName(context, action) {
  const myshape = await getShapeByType(
    context,
    action.sheet_position,
    Number(action.args[0]),
    Excel.ShapeType.image,
  );
  myshape.name = action.args[1].toString();
}

async function setPictureHeight(context, action) {
  const myshape = await getShapeByType(
    context,
    action.sheet_position,
    Number(action.args[0]),
    Excel.ShapeType.image,
  );
  myshape.height = Number(action.args[1]);
}

async function setPictureWidth(context, action) {
  const myshape = await getShapeByType(
    context,
    action.sheet_position,
    Number(action.args[0]),
    Excel.ShapeType.image,
  );
  myshape.width = Number(action.args[1]);
}

async function deletePicture(context, action) {
  const myshape = await getShapeByType(
    context,
    action.sheet_position,
    Number(action.args[0]),
    Excel.ShapeType.image,
  );
  myshape.delete();
}

async function addPicture(context, action) {
  const selectedAddress = await getSelectedRangeAddress(context);

  const imageBase64 = action["args"][0].toString();
  const colIndex = Number(action["args"][1]);
  const rowIndex = Number(action["args"][2]);
  let left = Number(action["args"][3]);
  let top = Number(action["args"][4]);

  const sheet = await getSheet(context, action);
  let anchorCell = sheet
    .getRangeByIndexes(rowIndex, colIndex, 1, 1)
    .load("left, top");
  await context.sync();
  left = Math.max(left, anchorCell.left);
  top = Math.max(top, anchorCell.top);
  const image = sheet.shapes.addImage(imageBase64);
  image.left = left;
  image.top = top;

  if (selectedAddress) {
    context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(selectedAddress)
      .select();
    await context.sync();
  }
}

async function updatePicture(context, action) {
  const selectedAddress = await getSelectedRangeAddress(context);

  const imageBase64 = action["args"][0].toString();
  const sheet = await getSheet(context, action);
  let image = await getShapeByType(
    context,
    action.sheet_position,
    Number(action.args[1]),
    Excel.ShapeType.image,
  );
  image = image.load("name, left, top, height, width");
  await context.sync();
  let imgName = image.name;
  let imgLeft = image.left;
  let imgTop = image.top;
  let imgHeight = image.height;
  let imgWidth = image.width;
  image.delete();

  const newImage = sheet.shapes.addImage(imageBase64);
  newImage.name = imgName;
  newImage.left = imgLeft;
  newImage.top = imgTop;
  newImage.height = imgHeight;
  newImage.width = imgWidth;

  if (selectedAddress) {
    context.workbook.worksheets
      .getActiveWorksheet()
      .getRange(selectedAddress)
      .select();
    await context.sync();
  }
}

async function alert(context, action) {
  let myPrompt = action.args[0].toString();
  let myTitle = action.args[1].toString();
  let myButtons = action.args[2].toString();
  let myMode = action.args[3].toString();
  let myCallback = action.args[4].toString();
  xlAlert(myPrompt, myTitle, myButtons, myMode, myCallback);
}

async function setRangeName(context, action) {
  let range = await getRange(context, action);
  context.workbook.names.add(action.args[0].toString(), range);
}

async function namesAdd(context, action) {
  let name = action.args[0].toString();
  let refersTo = action.args[1].toString();
  if (action.sheet_position == null) {
    context.workbook.names.add(name, refersTo);
  } else {
    let sheets = context.workbook.worksheets.load("items");
    await context.sync();
    sheets.items[action.sheet_position].names.add(name, refersTo);
  }
}

async function nameDelete(context, action) {
  let name = action.args[2].toString();
  let book_scope = Boolean(action.args[4]);
  let scope_sheet_index = Number(action.args[5]);
  if (book_scope === true) {
    context.workbook.names.getItem(name).delete();
  } else {
    let sheets = context.workbook.worksheets.load("items");
    await context.sync();
    sheets.items[scope_sheet_index].names.getItem(name).delete();
  }
}

async function runMacro(context, action) {
  await globalThis.callbacks[action.args[0].toString()](
    context,
    ...action.args.slice(1),
  );
}

async function rangeDelete(context, action) {
  let range = await getRange(context, action);
  let shift = action.args[0].toString();
  if (shift === "up") {
    range.delete(Excel.DeleteShiftDirection.up);
  } else if (shift === "left") {
    range.delete(Excel.DeleteShiftDirection.left);
  }
}

async function rangeInsert(context, action) {
  let range = await getRange(context, action);
  let shift = action.args[0].toString();
  if (shift === "down") {
    range.insert(Excel.InsertShiftDirection.down);
  } else if (shift === "right") {
    range.insert(Excel.InsertShiftDirection.right);
  }
}

async function rangeSelect(context, action) {
  let range = await getRange(context, action);
  range.select();
}

async function addTable(context, action) {
  let worksheets = context.workbook.worksheets.load("items");
  await context.sync();
  let mytable = worksheets.items[action.sheet_position].tables.add(
    action.args[0].toString(),
    Boolean(action.args[1]),
  );
  if (action.args[2] != null) {
    mytable.style = action.args[2].toString();
  }
  if (action.args[3] != null) {
    mytable.name = action.args[3].toString();
  }
}

async function setTableName(context, action) {
  const mytable = await getTable(context, action);
  mytable.name = action.args[1].toString();
}

async function resizeTable(context, action) {
  const mytable = await getTable(context, action);
  mytable.resize(action.args[1].toString());
}

async function showAutofilterTable(context, action) {
  const mytable = await getTable(context, action);
  mytable.showFilterButton = Boolean(action.args[1]);
}

async function showHeadersTable(context, action) {
  const mytable = await getTable(context, action);
  mytable.showHeaders = Boolean(action.args[1]);
}

async function showTotalsTable(context, action) {
  const mytable = await getTable(context, action);
  mytable.showTotals = Boolean(action.args[1]);
}

async function setTableStyle(context, action) {
  const mytable = await getTable(context, action);
  mytable.style = action.args[1].toString();
}

async function copyRange(context, action) {
  const destination = context.workbook.worksheets.items[
    parseInt(action.args[0].toString())
  ].getRange(action.args[1].toString());
  destination.copyFrom(await getRange(context, action));
}

async function copyFromRange(context, action) {
  const myRange = await getRange(context, action);
  const sourceRange = context.workbook.worksheets.items[
    parseInt(action.args[0].toString())
  ].getRange(action.args[1].toString());
  const copyType = action.args[2];
  const skipBlanks = Boolean(action.args[3]);
  const transpose = Boolean(action.args[4]);
  myRange.copyFrom(sourceRange, copyType, skipBlanks, transpose);
}

async function sheetDelete(context, action) {
  // TODO: use getSheet
  let worksheets = context.workbook.worksheets.load("items");
  await context.sync();
  worksheets.items[action.sheet_position].delete();
}

async function sheetClear(context, action) {
  // TODO: use getSheet
  let worksheets = context.workbook.worksheets.load("items");
  await context.sync();
  worksheets.items[action.sheet_position]
    .getRanges()
    .clear(Excel.ClearApplyTo.all);
}

async function sheetClearFormats(context, action) {
  // TODO: use getSheet
  let worksheets = context.workbook.worksheets.load("items");
  await context.sync();
  worksheets.items[action.sheet_position]
    .getRanges()
    .clear(Excel.ClearApplyTo.formats);
}

async function sheetClearContents(context, action) {
  // TODO: use getSheet
  let worksheets = context.workbook.worksheets.load("items");
  await context.sync();
  worksheets.items[action.sheet_position]
    .getRanges()
    .clear(Excel.ClearApplyTo.contents);
}

async function rangeGroup(context, action) {
  let myrange = await getRange(context, action);
  if (action.args[0].toString() == "columns") {
    myrange.group(Excel.GroupOption.byColumns);
  } else {
    myrange.group(Excel.GroupOption.byRows);
  }
}

async function rangeUngroup(context, action) {
  let myrange = await getRange(context, action);
  if (action.args[0].toString() == "columns") {
    myrange.ungroup(Excel.GroupOption.byColumns);
  } else {
    myrange.ungroup(Excel.GroupOption.byRows);
  }
}

async function freezePaneAtRange(context, action) {
  let sheet = await getSheet(context, action);
  let range = sheet.getRange(action.args[0].toString());
  sheet.freezePanes.freezeAt(range);
}

async function freezePaneUnfreeze(context, action) {
  let sheet = await getSheet(context, action);
  sheet.freezePanes.unfreeze();
}

async function rangeAdjustIndent(context, action) {
  let range = await getRange(context, action);
  range.format.adjustIndent(parseInt(action.args[0].toString()));
}
