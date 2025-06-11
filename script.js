let sheetCounter = 0;
let cellData = {};
let currentCell = null;

function createSpreadsheet(containerId) {
  const container = document.getElementById(containerId);
  const spreadsheet = document.createElement('div');
  spreadsheet.className = 'spreadsheet';
  spreadsheet.dataset.sheetId = sheetCounter;

  // Add column headers
  const emptyHeader = document.createElement('div');
  spreadsheet.appendChild(emptyHeader);

  for (let col = 0; col < 20; col++) {
    const colHeader = document.createElement('div');
    colHeader.className = 'header';
    colHeader.textContent = String.fromCharCode(65 + col);
    spreadsheet.appendChild(colHeader);
  }

  // Add rows and cells
  for (let row = 1; row <= 25; row++) {
    const rowHeader = document.createElement('div');
    rowHeader.className = 'row-header';
    rowHeader.textContent = row;
    spreadsheet.appendChild(rowHeader);

    for (let col = 0; col < 20; col++) {
      const cell = document.createElement('div');
      cell.className = 'cell';
      cell.contentEditable = true;

      const cellId = `${sheetCounter}_${String.fromCharCode(65 + col)}${row}`;
      cell.dataset.id = cellId;

      cell.addEventListener('click', () => currentCell = cell);
      cell.addEventListener('input', () => {
        cellData[cellId] = {
          value: cell.innerText,
          color: cell.style.color,
          backgroundColor: cell.style.backgroundColor,
          font: cell.style.fontFamily,
          size: cell.style.fontSize
        };
      });

      spreadsheet.appendChild(cell);
    }
  }

  container.appendChild(spreadsheet);
  currentCell = null;
  sheetCounter++;
}

function format(command) {
  document.execCommand(command, false, null);
}

function changeColor(color, type) {
  if (currentCell) {
    currentCell.style[type] = color;
    const cellId = currentCell.dataset.id;
    cellData[cellId] = cellData[cellId] || {};
    cellData[cellId][type === 'color' ? 'color' : 'backgroundColor'] = color;
  }
}

function changeFont(font) {
  if (currentCell) {
    currentCell.style.fontFamily = font;
  }
}

function changeFontSize(size) {
  if (currentCell) {
    currentCell.style.fontSize = `${10 + (size - 1) * 2}px`;
  }
}

function addNewSheet() {
  createSpreadsheet("sheet-container");
}

function exportToCSV() {
  let csv = "";
  for (let row = 1; row <= 25; row++) {
    let rowArr = [];
    for (let col = 0; col < 20; col++) {
      const cellId = `${sheetCounter - 1}_${String.fromCharCode(65 + col)}${row}`;
      const value = cellData[cellId]?.value || "";
      rowArr.push(`"${value.replace(/"/g, '""')}"`);
    }
    csv += rowArr.join(",") + "\n";
  }

  const blob = new Blob([csv], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "spreadsheet.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

createSpreadsheet("sheet-container");
