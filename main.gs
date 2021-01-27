function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
      .addItem('Convert selected range to Markdown Table', 'convetSeletedRangeToMarkdown')
      .addToUi();
}

function convetSeletedRangeToMarkdown() {
  const activeS = SpreadsheetApp.getActiveSheet();
  const selection = activeS.getSelection();
  const ranges = selection.getActiveRangeList().getRanges();
  const range = ranges[0];
  const valuesList = range.getValues();
  let md = '';
  valuesList.forEach((vl, i) => {
    let row = '|';
    if (i === 1) {
      vl.forEach((_, ii) => {
        row += ` --- |`;
      });
      md += `${row}\n`;
    }

    row = '|';
    vl.forEach((v, ii) => {
      let val = String(v).replace("\n", '<br>');
      row += ` ${val} |`;
    });
    md += `${row}\n`;
  });
  
  showDialog('Markdown Table', md);
}


function showDialog(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    title,
    message,
    ui.ButtonSet.OK);
}
