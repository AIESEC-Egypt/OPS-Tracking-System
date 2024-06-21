function AIActual() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("s2Dailynew");
  const sheetData = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
  const lcs = sheetData.flat(1);
  for (let i = 0; i < sprintStartDatenew.length; i++) {
    console.log(sprintStartDatenew[i]);

    for (let j = 0; j < 1; j++) {
      console.log(lcs[rowStartIndexnew[i] - 1 + j]);
      var url =
        "https://analytics.api.aiesec.org/v2/applications/analyze.json?access_token=" +
        "&start_date=" +
        sprintStartDatenew[i] +
        "&end_date=" +
        sprintEndDatenew[i] +
        "&performance_v3%5Boffice_id%5D=" +
        ai_codes[lcs[rowStartIndexnew[i] - 1 + j]];

      var response = UrlFetchApp.fetch(url, { method: "GET" }).getContentText();
      var data = JSON.parse(response);
      var list = [];
      list.push([
        data.open_icx.doc_count,
        data.open_i_programme_7.doc_count,
        data.open_i_programme_8.doc_count,
        data.open_i_programme_9.doc_count,
        data.open_ogx.doc_count,
        data.open_o_programme_7.doc_count,
        data.open_o_programme_8.doc_count,
        data.open_o_programme_9.doc_count,
        data.applied_total.applicants.value,
        data.i_applied_7.applicants.value,
        data.i_applied_8.applicants.value,
        data.i_applied_9.applicants.value,
        data.o_applied_7.applicants.value,
        data.o_applied_8.applicants.value,
        data.o_applied_9.applicants.value,
        data.matched_total.doc_count,
        data.i_matched_7.doc_count,
        data.i_matched_8.doc_count,
        data.i_matched_9.doc_count,
        data.o_matched_7.doc_count,
        data.o_matched_8.doc_count,
        data.o_matched_9.doc_count,
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
        data.finished_total.doc_count,
        data.i_finished_7.doc_count,
        data.i_finished_8.doc_count,
        data.i_finished_9.doc_count,
        data.o_finished_7.doc_count,
        data.o_finished_8.doc_count,
        data.o_finished_9.doc_count,
        data.completed_total.doc_count,
        data.i_completed_7.doc_count,
        data.i_completed_8.doc_count,
        data.i_completed_9.doc_count,
        data.o_completed_7.doc_count,
        data.o_completed_8.doc_count,
        data.o_completed_9.doc_count,
      ]);

      sheet
        .getRange(rowStartIndexnew[i] + j, 2, 1, list[0].length)
        .setValues(list);
    }
  }
}
