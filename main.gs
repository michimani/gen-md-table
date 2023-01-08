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
  let md = '';
  ranges.forEach((range, i) => {
    if (i > 0) {
      md += "\n\n";
    }
    const valuesList = range.getDisplayValues();
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
        let val = String(v).replace(/\n/g, '<br>');
        row += ` ${val} |`;
      });
      md += `${row}\n`;
    });
  })

  // showDialog('Markdown Table', md);
  showCopyDialog('Markdown Table', md);
}


function showDialog(title, message) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    title,
    message,
    ui.ButtonSet.OK);
}

function showCopyDialog(title, md) {
  var htmlOutput = HtmlService
    .createHtmlOutput(resultHtml(md));
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, title);
}

function resultHtml(md) {
  return `
  <style>
    #gen-md-result-copy-button {
      background-image: none;
      border: 1px solid transparent!important;
      border-radius: 4px;
      box-shadow: none;
      box-sizing: border-box;
      font-family: "Google Sans",Roboto,RobotoDraft,Helvetica,Arial,sans-serif;
      font-weight: 500;
      font-size: 14px;
      height: 36px;
      letter-spacing: 0.25px;
      line-height: 16px;
      padding: 9px 24px 11px 24px;
      background: white;
      border: 1px solid #dadce0!important;
      color: #188038;
      float: right;
  }
  #gen-md-result-copy-button:hover {
    cursor: pointer;
    background: #f0fff5;
  }
  #gen-md-result {
    resize: none;
    width: 100%;
    height: calc(100vh - 65px);
  }
  </style>
  <textarea id="gen-md-result" readonly>${md}</textarea>
  <button id="gen-md-result-copy-button" name="1" onclick="copyGenMdResult()">copy to clipboard</button>
  <script>
    function copyGenMdResult() {
      const ta = document.getElementById('gen-md-result');
      ta.focus();
      ta.select();
      document.execCommand('copy');
      google.script.host.close();
    }
  </script>
  `;
}
