document.getElementById("excelFile").addEventListener("change", handleFile, false);

let clientsData = {};

function handleFile(event) {
  const file = event.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });

    clientsData = {};

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      const headers = json[0];
      const monthIndexes = headers
        .map((h, i) => ({ h, i }))
        .filter(({ h }) =>
          h &&
          !["producto", "inventario", "venta"].includes(h.toString().toLowerCase())
        )
        .map(({ i }) => i);

      const inventoryIndex = headers.findIndex(h => h.toString().toLowerCase() === "inventario");
      const ventaIndex = headers.findIndex(h => h.toString().toLowerCase() === "venta");

      const rows = json.slice(1).filter(row => row.length > 0);
      let currentCategory = null;
      const categories = {};

      rows.forEach(row => {
        const nombreProducto = row[0];
        if (!nombreProducto || nombreProducto.trim() === "") return;

        if (!row[inventoryIndex] && !row[ventaIndex] && monthIndexes.every(i => !row[i])) {
          // Es una categoría
          currentCategory = nombreProducto.trim();
          if (!categories[currentCategory]) categories[currentCategory] = [];
        } else if (currentCategory) {
          // Es un producto
          const meses = monthIndexes.map(i => parseFloat(row[i]) || 0);
          const promedio = meses.length ? meses.reduce((a, b) => a + b, 0) / meses.length : 0;

          const inventario = parseFloat(row[inventoryIndex]) || 0;
          const venta = parseFloat(row[ventaIndex]) || 0;
          const invVenta = inventario + venta;

          categories[currentCategory].push({
            producto: nombreProducto,
            promedio,
            invVenta
          });
        }
      });

      // Calculamos totales por categoría
      const resumen = [];
      Object.entries(categories).forEach(([categoria, productos]) => {
        const sumaPromedios = productos.reduce((sum, p) => sum + p.promedio, 0);
        const sumaInvVenta = productos.reduce((sum, p) => sum + p.invVenta, 0);
        const division = sumaPromedios ? sumaInvVenta / sumaPromedios : 0;

        resumen.push({
          categoria,
          productos,
          sumaPromedios,
          sumaInvVenta,
          division
        });
      });

      clientsData[sheetName] = resumen;
    });

    renderClientButtons();
    showClient(Object.keys(clientsData)[0]);
  };

  reader.readAsArrayBuffer(file);
}

function renderClientButtons() {
  const container = document.getElementById("clientButtons");
  container.innerHTML = "";

  Object.keys(clientsData).forEach(clientName => {
    const btn = document.createElement("button");
    btn.textContent = clientName;
    btn.onclick = () => showClient(clientName);
    container.appendChild(btn);
  });
}

function showClient(clientName) {
  const container = document.getElementById("tablesContainer");
  container.innerHTML = "";

  const resumen = clientsData[clientName];
  const tableDiv = document.createElement("div");
  tableDiv.className = "table-container active";

  resumen.forEach(categoria => {
    const table = document.createElement("table");

    // Título categoría
    const title = document.createElement("caption");
    title.textContent = categoria.categoria;
    title.style.fontWeight = "bold";
    title.style.fontSize = "1.2em";
    table.appendChild(title);

    // Cabecera
    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");
    ["Producto", "Promedio Meses", "Inventario + Venta"].forEach(h => {
      const th = document.createElement("th");
      th.textContent = h;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Datos productos
    const tbody = document.createElement("tbody");
    categoria.productos.forEach(prod => {
      const row = document.createElement("tr");
      [prod.producto, prod.promedio.toFixed(1), prod.invVenta.toFixed(1)].forEach(text => {
        const td = document.createElement("td");
        td.textContent = text;
        row.appendChild(td);
      });
      tbody.appendChild(row);
    });

    // Fila resumen categoría
    const resumenRow = document.createElement("tr");
    resumenRow.style.fontWeight = "bold";

    const td1 = document.createElement("td");
    td1.textContent = "RESUMEN";
    const td2 = document.createElement("td");
    td2.textContent = categoria.sumaPromedios.toFixed(1);
    const td3 = document.createElement("td");
    td3.textContent = categoria.sumaInvVenta.toFixed(1);

    resumenRow.appendChild(td1);
    resumenRow.appendChild(td2);
    resumenRow.appendChild(td3);
    tbody.appendChild(resumenRow);

    // Fila división
    const divisionRow = document.createElement("tr");
    const tdCat = document.createElement("td");
    tdCat.textContent = categoria.categoria;
    tdCat.colSpan = 2;

    const tdDivision = document.createElement("td");
    tdDivision.textContent = `Resultado: ${categoria.division.toFixed(1)}`;
    divisionRow.appendChild(tdCat);
    divisionRow.appendChild(tdDivision);
    tbody.appendChild(divisionRow);

    table.appendChild(tbody);
    tableDiv.appendChild(table);
  });

  container.appendChild(tableDiv);
}
function showClient(clientName) {
  const container = document.getElementById("tablesContainer");

  // Animación de salida
  container.classList.remove("show");
  container.classList.add("hide");

  setTimeout(() => {
    container.innerHTML = ""; // Limpiar contenido anterior
    document.getElementById("clientTitle").textContent = `Cliente: ${clientName}`;
    const resumen = clientsData[clientName];
    const tableDiv = document.createElement("div");
    tableDiv.className = "table-container";

    resumen.forEach(categoria => {
      const table = document.createElement("table");

      const title = document.createElement("caption");
      title.textContent = categoria.categoria;
      table.appendChild(title);

      const thead = document.createElement("thead");
      const headerRow = document.createElement("tr");
      ["Producto", "Promedio Meses", "Inventario + Venta"].forEach(h => {
        const th = document.createElement("th");
        th.textContent = h;
        headerRow.appendChild(th);
      });
      thead.appendChild(headerRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      categoria.productos.forEach(prod => {
        const row = document.createElement("tr");
        [prod.producto, prod.promedio.toFixed(1), prod.invVenta.toFixed(1)].forEach(text => {
          const td = document.createElement("td");
          td.textContent = text;
          row.appendChild(td);
        });
        tbody.appendChild(row);
      });

      const resumenRow = document.createElement("tr");
      const td1 = document.createElement("td");
      td1.textContent = "RESUMEN";
      const td2 = document.createElement("td");
      td2.textContent = categoria.sumaPromedios.toFixed(1);
      const td3 = document.createElement("td");
      td3.textContent = categoria.sumaInvVenta.toFixed(1);
      resumenRow.appendChild(td1);
      resumenRow.appendChild(td2);
      resumenRow.appendChild(td3);
      resumenRow.style.fontWeight = "bold";
      tbody.appendChild(resumenRow);

      const divisionRow = document.createElement("tr");
      const tdCat = document.createElement("td");
      tdCat.colSpan = 2;
      tdCat.textContent = categoria.categoria;
      const tdDivision = document.createElement("td");
      tdDivision.textContent = `Resultado: ${categoria.division.toFixed(1)}`;
      divisionRow.appendChild(tdCat);
      divisionRow.appendChild(tdDivision);
      tbody.appendChild(divisionRow);

      table.appendChild(tbody);
      tableDiv.appendChild(table);
    });

    container.appendChild(tableDiv);

    // Animación de entrada
    setTimeout(() => {
      tableDiv.classList.add("show");
    }, 50);
  }, 200);
}
