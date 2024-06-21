function GlobalActual() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "Base actuals term 23_AI"
  );
  const sheetData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const lcs = sheetData.flat(1);
  for (let i = 0; i < sprintStartDateglobal.length; i++) {
    console.log(sprintStartDateglobal[i]);

    for (let j = 0; j < 144; j++) {
      console.log(lcs[rowStartIndexglobal[i] - 1 + j]);
      var url =
        "https://analytics.api.aiesec.org/v2/applications/analyze.json?access_token=" +
        "&start_date=" +
        sprintStartDateglobal[i] +
        "&end_date=" +
        sprintEndDateglobal[i] +
        "&performance_v3%5Boffice_id%5D=" +
        global_codes[lcs[rowStartIndexglobal[i] - 1 + j]];

      var response = UrlFetchApp.fetch(url, { method: "GET" }).getContentText();
      var data = JSON.parse(response);
      var list = [];
      list.push([
        data.applied_total.applicants.value,
        data.i_applied_7.applicants.value,
        data.i_applied_8.applicants.value,
        data.i_applied_9.applicants.value,
        data.o_applied_7.applicants.value,
        data.o_applied_8.applicants.value,
        data.o_applied_9.applicants.value,
        data.approved_total.doc_count,
        data.i_approved_7.doc_count,
        data.i_approved_8.doc_count,
        data.i_approved_9.doc_count,
        data.o_approved_7.doc_count,
        data.o_approved_8.doc_count,
        data.o_approved_9.doc_count,
        data.realized_total.doc_count,
        data.i_realized_7.doc_count,
        data.i_realized_8.doc_count,
        data.i_realized_9.doc_count,
        data.o_realized_7.doc_count,
        data.o_realized_8.doc_count,
        data.o_realized_9.doc_count,
        data.completed_total.doc_count,
        data.i_completed_7.doc_count,
        data.i_completed_8.doc_count,
        data.i_completed_9.doc_count,
        data.o_completed_7.doc_count,
        data.o_completed_8.doc_count,
        data.o_completed_9.doc_count,
      ]);

      sheet
        .getRange(rowStartIndexglobal[i] + j, 2, 1, list[0].length)
        .setValues(list);
    }
  }
}
