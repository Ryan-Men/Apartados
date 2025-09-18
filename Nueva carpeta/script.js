let workbookData = [];
let headers = [];
const hiddenColumn = "Estado"; // columna oculta

document.getElementById("importBtn").addEventListener("click", () => {
  document.getElementById("fileInput").click();
});

document.getElementById("fileInput").addEventListener("change", handleFile);
document.getElementById("exportBtn").addEventListener("click", exportExcel);
document.getElementById("addRowBtn").addEventListener("click", addRow);
document.getElementById("searchInput").addEventListener("input", filterTable);

function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  const isJson = file.name.endsWith(".json");

  reader.onload = (event) => {
    if (isJson) {
      workbookData = JSON.parse(event.target.result);
      headers = Object.keys(workbookData[0]);
    } else {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      let rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false });
      headers = rows[0];

      // Asegurarse de tener la columna Estado
      if(!headers.includes(hiddenColumn)) headers.push(hiddenColumn);

      workbookData = rows.slice(1).map(r => {
        const obj = {};
        headers.forEach((h,i)=> obj[h] = r[i] || "");
        if(!obj[hiddenColumn]) obj[hiddenColumn] = "blanco"; // predeterminado
        return obj;
      });
    }
    renderTable();
  };

  if(isJson) reader.readAsText(file);
  else reader.readAsArrayBuffer(file);
}

function renderTable() {
  const container = document.getElementById("tableContainer");
  container.innerHTML = "";

  const table = document.createElement("table");
  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  headRow.appendChild(document.createElement("th")); // columna de botones

  headers.forEach(h => {
    if(h !== hiddenColumn){
      const th = document.createElement("th");
      th.textContent = h;
      headRow.appendChild(th);
    }
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");

  workbookData.forEach((rowData) => {
    const row = document.createElement("tr");

    // Columna botón
    const tdBtn = document.createElement("td");
    const btn = document.createElement("button");
    btn.classList.add("status-btn");

    // Inicializar botón con el estado que viene del archivo
    updateButtonAndRow(btn, rowData, row, rowData[hiddenColumn]);

    btn.addEventListener("click", () => {
      // Cambia estado en ciclo: blanco → amarillo → verde → rojo → blanco
      const current = rowData[hiddenColumn].toLowerCase();
      let next;
      if(current==="blanco") next="amarillo";
      else if(current==="amarillo") next="verde";
      else if(current==="verde") next="rojo";
      else next="blanco";
      rowData[hiddenColumn] = next;
      updateButtonAndRow(btn, rowData, row, next);
    });

    tdBtn.appendChild(btn);
    row.appendChild(tdBtn);

    // Celdas editables
    headers.forEach(h => {
      if(h!==hiddenColumn){
        const td = document.createElement("td");
        td.textContent = rowData[h] || "";
        td.setAttribute("contenteditable","true");
        td.addEventListener("input", e => rowData[h] = e.target.textContent);
        row.appendChild(td);
      }
    });

    tbody.appendChild(row);
  });

  table.appendChild(tbody);
  container.appendChild(table);
}

function updateButtonAndRow(btn, rowData, row, estado){
  btn.className="status-btn";
  estado = estado.toLowerCase();

  if(estado==="verde"){
    btn.classList.add("status-verde"); btn.textContent="✔"; row.style.backgroundColor="green";
  } else if(estado==="amarillo"){
    btn.classList.add("status-amarillo"); btn.textContent="⚠"; row.style.backgroundColor="yellow";
  } else if(estado==="rojo"){
    btn.classList.add("status-rojo"); btn.textContent="✖"; row.style.backgroundColor="red";
  } else if(estado==="blanco"){
    btn.classList.add("status-blanco"); btn.textContent="⚪"; row.style.backgroundColor="white";
  }
}

function addRow() {
  const newRow = {};
  headers.forEach(h=>newRow[h]="");
  newRow[hiddenColumn]="blanco";
  workbookData.push(newRow);
  renderTable();
}

function exportExcel() {
  if(!workbookData.length) return;
  const wb = XLSX.utils.book_new();
  const ws_data = [];

  // Incluir headers completos, incluyendo Estado
  if(!headers.includes(hiddenColumn)) headers.push(hiddenColumn);
  ws_data.push(headers);

  // Incluir cada fila con el estado
  workbookData.forEach(row=>{
    const rowArr = headers.map(h=>row[h]||"");
    ws_data.push(rowArr);
  });

  const ws = XLSX.utils.aoa_to_sheet(ws_data);
  XLSX.utils.book_append_sheet(wb, ws, "Apartados");
  XLSX.writeFile(wb, "apartados_exportados.xlsx");
}

function filterTable() {
  const filter = document.getElementById("searchInput").value.toLowerCase();
  const table = document.querySelector("table tbody");
  table.querySelectorAll("tr").forEach(tr=>{
    const text = Array.from(tr.querySelectorAll("td[contenteditable=true]")).map(td=>td.textContent.toLowerCase()).join(" ");
    tr.style.display = text.includes(filter) ? "" : "none";
  });
}
