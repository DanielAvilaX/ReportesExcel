document.getElementById("excelFile").addEventListener("change", handleFile, false);

let clientsData = {};
let headersLocal = [];
let monthIndexes = [];
let localHeadersPorCliente = {};
let localMonthIndexesPorCliente = {};

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

      const localHeaders = json[0];

      const rawMonthVariants = [
        "enero", "ene", "febrero", "feb", "marzo", "mar", "abril", "abr", "mayo", "may",
        "junio", "jun", "julio", "jul", "agosto", "ago", "septiembre", "sep", "set",
        "octubre", "oct", "noviembre", "nov", "diciembre", "dic",
        "january", "jan", "february", "feb", "march", "mar", "april", "apr",
        "may", "june", "jun", "july", "jul", "august", "aug", "september", "sep",
        "october", "oct", "november", "nov", "december", "dec"
      ];

      const monthNames = rawMonthVariants.map(m =>
        m.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase()
      );

      const localMonthIndexes = localHeaders
        .map((h, i) => {
          const normalized = h?.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
          return { h: normalized, i };
        })
        .filter(({ h }) => monthNames.includes(h))
        .map(({ i }) => i);

      localHeadersPorCliente[sheetName] = localHeaders;
      localMonthIndexesPorCliente[sheetName] = localMonthIndexes;

      const inventoryIndex = localHeaders.findIndex(h => h.toString().toLowerCase() === "inventario");
      const ventaIndex = localHeaders.findIndex(h => h.toString().toLowerCase() === "venta");

      const rows = json.slice(1).filter(row => row.length > 0);
      let currentCategory = null;
      const categories = {};

      rows.forEach(row => {
        const nombreProducto = row[0];
        if (!nombreProducto || nombreProducto.trim() === "") return;

        if (!row[inventoryIndex] && !row[ventaIndex] && localMonthIndexes.every(i => !row[i])) {
          currentCategory = nombreProducto.trim();
          if (!categories[currentCategory]) categories[currentCategory] = [];
        } else if (currentCategory) {
          const meses = localMonthIndexes.map(i => parseFloat(row[i]) || 0);
          const promedio = meses.length ? meses.reduce((a, b) => a + b, 0) / meses.length : 0;
          const inventario = parseFloat(row[inventoryIndex]) || 0;
          const venta = parseFloat(row[ventaIndex]) || 0;
          const invVenta = inventario + venta;

          categories[currentCategory].push({
            producto: nombreProducto,
            promedio,
            invVenta,
            meses,
            etiquetas: localMonthIndexes.map(i => localHeaders[i]),
            cliente: sheetName,

          });
        }
      });

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

    const exportBtn = document.getElementById("btnExportarExcel");
    exportBtn.style.pointerEvents = "auto";
    exportBtn.style.opacity = "1";
    exportBtn.onclick = exportarTodoExcel;
    document.getElementById("expandirGlobalWrapper").style.display = "block";
    document.getElementById("expandirGlobalWrapperBottom").style.display = "block";

  };

  reader.readAsArrayBuffer(file);
  document.getElementById("expandirGlobalWrapper").style.display = "block";
}

function renderClientButtons() {
  const container = document.getElementById("clientButtons");
  container.innerHTML = "";

  Object.keys(clientsData).forEach(clientName => {
    const btn = document.createElement("button");
    btn.textContent = clientName;
    btn.onclick = () => showClient(clientName);
    container.appendChild(btn); // <- ahora sí correcto
  });
}


function showClient(clientName) {
  const container = document.getElementById("tablesContainer");
  container.innerHTML = "";
  const selectedClientNameElement = document.getElementById("selectedClientName");
  selectedClientNameElement.textContent = `Cliente: ${clientName}`;
  selectedClientNameElement.style.color = "white";
  selectedClientNameElement.style.textShadow = "0 0 5px #000000, 0 0 10px #000000, 0 0 20px #000000";

  const resumen = clientsData[clientName];

  resumen.forEach((categoria, index) => {
    const categoriaBox = document.createElement("div");
    categoriaBox.className = "categoria-container"; // ✅ aquí va
    categoriaBox.style.marginBottom = "20px";

    const resumenHeader = document.createElement("div");
    resumenHeader.className = "categoria-header";

    const titulo = document.createElement("span");
    titulo.textContent = categoria.categoria;
    titulo.className = "categoria-titulo";

    const resultado = document.createElement("span");
    resultado.textContent = `Resultado: ${categoria.division.toFixed(1)}`;
    resultado.className = "categoria-resultado";

    const toggleBtn = document.createElement("button");
    toggleBtn.className = "toggle-btn";
    toggleBtn.textContent = "Expandir tabla";
    toggleBtn.onclick = () => toggleTabla(index, toggleBtn);

    resumenHeader.appendChild(titulo);
    resumenHeader.appendChild(resultado);
    resumenHeader.appendChild(toggleBtn);
    categoriaBox.appendChild(resumenHeader);

    const collapsible = document.createElement("div");
    collapsible.className = "tabla-collapsable";
    collapsible.id = `collapsable-${index}`;

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
    resumenRow.style.fontWeight = "bold";
    resumenRow.appendChild(createCell("RESUMEN"));
    resumenRow.appendChild(createCell(categoria.sumaPromedios.toFixed(1)));
    resumenRow.appendChild(createCell(categoria.sumaInvVenta.toFixed(1)));
    tbody.appendChild(resumenRow);

    table.appendChild(tbody);
    collapsible.appendChild(table);

    const accionesRow = document.createElement("div");
    accionesRow.style.display = "flex";
    accionesRow.style.justifyContent = "space-between";
    accionesRow.style.alignItems = "center";
    accionesRow.style.margin = "10px 20px";

    const btnGrafica = document.createElement("button");
    btnGrafica.textContent = "Grafica";
    btnGrafica.className = "toggle-btn";
    btnGrafica.onclick = () => mostrarGraficaCategoria(categoria);

    const btnContraer = document.createElement("button");
    btnContraer.textContent = "Contraer tabla";
    btnContraer.className = "toggle-btn";
    btnContraer.onclick = () => toggleTabla(index, toggleBtn);

    accionesRow.appendChild(btnGrafica);
    accionesRow.appendChild(btnContraer);
    collapsible.appendChild(accionesRow);

    categoriaBox.appendChild(collapsible);
    container.appendChild(categoriaBox); // ✅
  });
}


function createCell(content) {
  const td = document.createElement("td");
  td.textContent = content;
  return td;
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
        data.push([prod.producto, prod.promedio.toFixed(1), prod.invVenta.toFixed(1)]);
      });

      data.push(["RESUMEN", categoria.sumaPromedios.toFixed(1), categoria.sumaInvVenta.toFixed(1)]);
      data.push([categoria.categoria, "", `Resultado: ${categoria.division.toFixed(1)}`]);
      data.push([]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, cliente);
  }

  XLSX.writeFile(wb, "ResumenClientes.xlsx");
}
let chartInstance = null;

function mostrarGraficaCategoria(categoria) {
  const labelsSet = new Set();
  categoria.productos.forEach(producto => {
    producto.etiquetas.forEach(etiqueta => labelsSet.add(etiqueta));
  });
  const labels = Array.from(labelsSet);

  const datasets = categoria.productos.map(prod => {
    const dataMap = {};
    prod.etiquetas.forEach((et, i) => {
      dataMap[et] = prod.meses[i];
    });

    return {
      label: prod.producto,
      data: labels.map(et => dataMap[et] ?? null), // Evita valores cruzados
      borderColor: getRandomColor(),
      fill: false
    };
  });


  const ctx = document.getElementById("modalChartCanvas").getContext("2d");

  // Destruir instancia previa
  if (window.currentChartInstance) {
    window.currentChartInstance.destroy();
  }



  window.currentChartInstance = new Chart(ctx, {

    type: 'line',
    data: {
      labels,
      datasets
    },
    options: {
      responsive: true,
      plugins: {
        title: {
          display: true,
          text: `Gráfico de ${categoria.categoria}`,
          font: {
            size: 20
          }
        }
      }
    }
  });

  document.getElementById("chartModal").style.display = "flex";
}

function closeModal() {
  document.getElementById("chartModal").style.display = "none";
}

function getRandomColor() {
  return `hsl(${Math.floor(Math.random() * 360)}, 70%, 50%)`;
}
function toggleTabla(index, btn) {
  const section = document.getElementById(`collapsable-${index}`);
  const expanded = section.classList.toggle("expandida");
  btn.textContent = expanded ? "Contraer tabla" : "Expandir tabla";
}
const btnGlobalTop = document.getElementById("toggleAllTables");
const btnGlobalBottom = document.getElementById("toggleAllTablesBottom");

btnGlobalTop.addEventListener("click", () => toggleTodasLasTablas());
btnGlobalBottom.addEventListener("click", () => toggleTodasLasTablas());

function toggleTodasLasTablas() {
  const allCollapsibles = document.querySelectorAll(".tabla-collapsable");
  const expandir = ![...allCollapsibles].some(el => el.classList.contains("expandida"));

  document.querySelectorAll(".categoria-header").forEach((header, i) => {
    const toggleBtn = header.querySelector(".toggle-btn");
    const collapsible = document.getElementById(`collapsable-${i}`);
    if (expandir) {
      collapsible.classList.add("expandida");
      toggleBtn.textContent = "Contraer tabla";
    } else {
      collapsible.classList.remove("expandida");
      toggleBtn.textContent = "Expandir tabla";
    }
  });

  // Cambiar texto en ambos botones globales
  btnGlobalTop.textContent = expandir ? "Contraer todas las tablas" : "Expandir todas las tablas";
  btnGlobalBottom.textContent = expandir ? "Contraer todas las tablas" : "Expandir todas las tablas";
}


