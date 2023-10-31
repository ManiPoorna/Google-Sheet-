
const totalRows = 100;
const totalColumns = 26;

// Table Header
const tableHeadRow = document.getElementById("table-header");
// Function to add header data to the table
const addHeaderToTable = () => {
  for (let i = 0; i < totalColumns; i++) {
    const column = document.createElement("th");
    column.innerText = String.fromCharCode(i + 65);
    tableHeadRow.append(column)
  }
}
addHeaderToTable();

// Table Body
const tableBody = document.getElementById("table-body");
// Function to add body and body cells data to the table
const addBodyToTable = () => {
  let columns = 26;
  let rows = 100;
  for (let row = 0; row < totalRows; row++) {
    //creating tr 
    const tr = document.createElement("tr");
    // crating th for left side number header
    const th = document.createElement("th");
    th.innerText = row + 1;
    tr.append(th);
    for (let col = 0; col < totalColumns; col++) {
      // adding cells to tr created above
      const td = document.createElement("td");
      td.setAttribute("contenteditable", true)
      let char = String.fromCharCode(col + 65);
      td.setAttribute("id", `${char}${row + 1}`)
      td.addEventListener("click", (e) => handleClickOnCell(e))
      td.addEventListener("input", (e) => updateMatrix(e.target))
      tr.append(td);
    }
    tableBody.append(tr)
  }
}
addBodyToTable();

let currentCell = "";
function handleClickOnCell(e) {
  // console.log(e)
  selectedCell = e.target;
  currentCell = selectedCell
  document.getElementById("current-cell").innerText = selectedCell.id
}

// accessing all buttons like cut, copy, paste, etc..
const copyBtn = document.getElementById("copy-btn")
const cutBtn = document.getElementById("cut-btn")
const pasteBtn = document.getElementById("paste-btn")
const boldBtn = document.getElementById("bold-btn")
const italicBtn = document.getElementById("italic-btn")
const underlineBtn = document.getElementById("underline-btn")
const alignLeftBtn = document.getElementById("align-left-btn")
const alignCenterBtn = document.getElementById("align-center-btn")
const alignRightBtn = document.getElementById("align-right-btn")
const selectSizeBtn = document.getElementById("font-size");
const selectStyleBtn = document.getElementById("font-style");
const bgColorBtn = document.getElementById("bg-color-btn");
const fontColorBtn = document.getElementById("font-color-btn");
const downloadBtn = document.getElementById("download-btn");
const uploadBtn = document.getElementById("upload-btn");

let copiedValue = "";
copyBtn.addEventListener("click", (e) => {
  copiedValue = currentCell.innerText
  updateMatrix(currentCell)
})
cutBtn.addEventListener("click", (e) => {
  copiedValue = currentCell.innerText;
  currentCell.innerText = "";
})
pasteBtn.addEventListener("click", (e) => {
  currentCell.innerText += copiedValue
  updateMatrix(currentCell)
})
boldBtn.addEventListener("click", (e) => {
  if (currentCell.style.fontWeight === "bold") {
    currentCell.style.fontWeight = "normal";
    updateMatrix(currentCell)
    return;
  }
  currentCell.style.fontWeight = "bold";
  updateMatrix(currentCell)
})
italicBtn.addEventListener("click", (e) => {
  if (currentCell.style.fontStyle === "italic") {
    currentCell.style.fontStyle = "normal";
    updateMatrix(currentCell)
    return;
  }
  currentCell.style.fontStyle = "italic";
  updateMatrix(currentCell)
})
underlineBtn.addEventListener("click", (e) => {
  if (currentCell.style.textDecoration === "underline") {
    currentCell.style.textDecoration = "none";
    updateMatrix(currentCell)
    return;
  }
  currentCell.style.textDecoration = "underline";
  updateMatrix(currentCell)
})
alignLeftBtn.addEventListener("click", (e) => {
  currentCell.style.textAlign = "left";
  updateMatrix(currentCell)
})
alignRightBtn.addEventListener("click", (e) => {
  currentCell.style.textAlign = "right";
  updateMatrix(currentCell)
})
alignCenterBtn.addEventListener("click", (e) => {
  currentCell.style.textAlign = "center";
  updateMatrix(currentCell)
})
selectSizeBtn.addEventListener("change", (e) => {
  if (currentCell.innerText) {
    currentCell.style.fontSize = `${e.target.value}px`;
    updateMatrix(currentCell)
  }
  updateMatrix(currentCell)
})
selectStyleBtn.addEventListener('change', (e) => {
  currentCell.style.fontFamily = `${e.target.value}`;
  updateMatrix(currentCell)
})
bgColorBtn.addEventListener('input', (e) => {
  currentCell.style.backgroundColor = e.target.value
  updateMatrix(currentCell)
})
fontColorBtn.addEventListener('input', (e) => {
  currentCell.style.color = e.target.value;
  updateMatrix(currentCell)
})
downloadBtn.addEventListener("click", (e) => { 
  console.log('first')
  const excelSheetData = JSON.stringify(matrix);
  const blob = new Blob([excelSheetData],{type: "application/json"})
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob);
  link.download = "Excel Sheet Data.json";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
})

uploadBtn.addEventListener('change', (e) => {
  let file = e.target.files[0];
  if (file) {
    let reader = new FileReader();
    reader.readAsText(file)
    reader.onload = function(e){
      // it will read the array
      let fileContent = e.target.result; 
      try {
        matrix = JSON.parse(fileContent);
        // addign content to the sheet
        matrix.forEach((row) => {
          row.forach((cell) => {
            if (cell.id) {
              let editedCell = document.getElementById(cell.id);
              editedCell.innerText = cell.text;
              editedCell.style = cell.style
            }
          })
        })
      }
      catch (error) {
        console.log(error)
      }
    }
  }
})


// Creating array for storing Data from Excel Sheet
let matrix = new Array(totalRows);

for (let row = 0; row < totalRows; row++) {
  matrix[row] = new Array(totalColumns);
  for (let col = 0; col < totalColumns; col++) {
    matrix[row][col] = {};
  }
}


function updateMatrix(currentCell) {
  let j = currentCell.id[0].charCodeAt(0) - 65;
  let i = parseInt(currentCell.id.substring(1)) - 1;
  let obj = {
    text: currentCell.innerText,
    style: currentCell.style.cssText,
    id: currentCell.id
  }
  matrix[i][j] = obj
}
// console.log(matrix);