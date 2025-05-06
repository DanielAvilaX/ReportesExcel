document.getElementById("excelFile").addEventListener("change", handleFile, false);

let clientsData = {};
let chartInstance = null;
let headers = [];
let monthIndexes = [];
let headersGlobal = [];
let monthIndexesGlobal = [];




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
      // Detectar columnas de meses válidos
      const rawMonthVariants = [
        // Español - completo y abreviado
        "enero", "ene",
        "febrero", "feb",
        "marzo", "mar",
        "abril", "abr",
        "mayo", "may",
        "junio", "jun",
        "julio", "jul",
        "agosto", "ago",
        "septiembre", "sep", "set", // setiembre común en LATAM
        "octubre", "oct",
        "noviembre", "nov",
        "diciembre", "dic",

        // Inglés - completo y abreviado
        "january", "jan",
        "february", "feb",
        "march", "mar",
        "april", "apr",
        "may", // igual en ambos idiomas
        "june", "jun",
        "july", "jul",
        "august", "aug",
        "september", "sep",
        "october", "oct",
        "november", "nov",
        "december", "dec"
      ];

      // Preparamos la lista final de coincidencias en minúsculas y sin tildes
      const monthNames = rawMonthVariants.map(m =>
        m.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
      );


      const monthIndexes = headers
        .map((h, i) => {
          const normalized = h?.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
          return { h: normalized, i };
        })
        .filter(({ h }) => monthNames.includes(h))
        .map(({ i }) => i);

      headersGlobal = headers;
      monthIndexesGlobal = monthIndexes;



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
            invVenta,
            _row: row // Guardamos toda la fila original
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
    // Crear o reactivar botón de exportar
    const exportBtn = document.getElementById("btnExportarExcel");

    exportBtn.style.pointerEvents = "auto";
    exportBtn.style.opacity = "1";
    exportBtn.onclick = exportarTodoExcel;



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
  document.getElementById("selectedClientName").textContent = `Cliente: ${clientName}`;
  document.getElementById("chartCanvas").style.display = "none"; // ocultar canvas


  const resumen = clientsData[clientName];
  const tableDiv = document.createElement("div");
  tableDiv.className = "table-container show";

  resumen.forEach(categoria => {
    // Contenedor de botones alineados horizontalmente
    const btnContainer = document.createElement("div");
    btnContainer.style.display = "flex";
    btnContainer.style.flexWrap = "wrap";
    btnContainer.style.justifyContent = "center";
    btnContainer.style.gap = "10px";
    btnContainer.style.marginBottom = "20px";

    categoria.productos.forEach(prod => {
      const btn = document.createElement("button");
      btn.textContent = `Ver gráfica de ${prod.producto}`;
      btn.style.padding = "8px";
      btn.style.background = "#4caf50";
      btn.style.color = "#fff";
      btn.style.border = "none";
      btn.style.borderRadius = "4px";
      btn.style.cursor = "pointer";
      btn.onmouseover = () => btn.style.background = "#388e3c";
      btn.onmouseleave = () => btn.style.background = "#4caf50";

      btn.onclick = () => renderChartProducto(prod.producto, monthIndexesGlobal, prod._row, headersGlobal);
      btnContainer.appendChild(btn);
    });

    tableDiv.appendChild(btnContainer);


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
  const allProducts = resumen.flatMap(cat => cat.productos);
  const btnTodos = document.createElement("button");
  btnTodos.textContent = "📊 Ver gráfica de todos los productos";
  btnTodos.style.marginTop = "20px";
  btnTodos.style.padding = "10px";
  btnTodos.style.background = "#ff9800";
  btnTodos.style.color = "#fff";
  btnTodos.style.border = "none";
  btnTodos.style.borderRadius = "4px";
  btnTodos.style.cursor = "pointer";
  btnTodos.onmouseover = () => btnTodos.style.background = "#e68900";
  btnTodos.onmouseleave = () => btnTodos.style.background = "#ff9800";
  btnTodos.onclick = () => renderChartTodosProductos(allProducts, monthIndexesGlobal, headersGlobal);
  tableDiv.appendChild(btnTodos);

}
function renderChartProducto(productName, monthIndexes, row, headers) {
  const ctx = document.getElementById("modalChartCanvas").getContext("2d");
  if (window.currentChart) {
    window.currentChart.destroy();
  }

  const labels = monthIndexes.map(i => headers[i]);
  const data = monthIndexes.map(i => parseFloat(row[i]));

  window.currentChart = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: productName,
        data,
        borderColor: 'rgba(75, 192, 192, 1)',
        fill: false,
        tension: 0.1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });

  openModal();
}

function renderChartTodosProductos(productos, monthIndexes, headers) {
  const ctx = document.getElementById("chartCanvas").getContext("2d");
  const labels = monthIndexes.map(i => headers[i]);

  const datasets = productos.map(prod => ({
    label: prod.producto,
    data: monthIndexes.map(i => parseFloat(prod._row[i]) || 0),
    fill: false,
    borderColor: getRandomColor(),
    tension: 0.2
  }));

  if (chartInstance) chartInstance.destroy();

  chartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets
    },
    options: {
      responsive: true,
      animation: {
        duration: 800
      },
      plugins: {
        title: {
          display: true,
          text: 'Ventas de todos los productos'
        }
      }
    }
  });

  document.getElementById("chartCanvas").style.display = "block";
}
function getRandomColor() {
  const letters = '0123456789ABCDEF';
  let color = '#';
  for (let i = 0; i < 6; i++) color += letters[Math.floor(Math.random() * 16)];
  return color;
}
function openModal() {
  document.getElementById("chartModal").style.display = "flex";
}

function closeModal() {
  document.getElementById("chartModal").style.display = "none";
}

function exportarTodoExcel() {
  if (!clientsData || Object.keys(clientsData).length === 0) {
    alert("No hay datos para exportar. Por favor sube primero el archivo.");
    return;
  }
  const wb = XLSX.utils.book_new();

  for (const cliente in clientsData) {
    const resumen = clientsData[cliente];

    const data = [];

    resumen.forEach(categoria => {
      data.push([`Categoría: ${categoria.categoria}`]);
      data.push(["Producto", "Promedio Meses", "Inventario + Venta"]);

      categoria.productos.forEach(prod => {
        data.push([
          prod.producto,
          prod.promedio.toFixed(1),
          prod.invVenta.toFixed(1)
        ]);
      });

      // Fila resumen
      data.push([
        "RESUMEN",
        categoria.sumaPromedios.toFixed(1),
        categoria.sumaInvVenta.toFixed(1)
      ]);

      // División
      data.push([
        categoria.categoria,
        "",
        `Resultado: ${categoria.division.toFixed(1)}`
      ]);

      // Línea vacía entre categorías
      data.push([]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, cliente);
  }

  XLSX.writeFile(wb, "ResumenClientes.xlsx");
}
