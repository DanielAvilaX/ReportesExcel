document.getElementById("excelFile").addEventListener("change", handleFile, false);

let clientsData = {};

function handleFile(event) {
  const file = event.target.files[0];
  if (!file || !file.name.match(/\.(xlsx|xls)$/)) {
    alert("Por favor selecciona un archivo Excel válido.");
    return;
  }

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    clientsData = {};

    workbook.SheetNames.forEach(sheetName => {
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      if (!json || json.length < 17) return;

      const etiquetas = json[10].slice(12, 18).map(e => e?.toString().trim());

      const dataCliente = {};
      let negocioActual = null;
      let lineaActual = null;

      for (let i = 14; i < json.length; i++) {
        const row = json[i];
        const celdaC = row[2]?.toString().trim();

        const esNegocio = /^\d{3}\s*-\s*/.test(celdaC);
        const esLinea = celdaC?.toLowerCase().startsWith("línea:");

        if (esNegocio && json[i + 1]?.[2]?.toLowerCase().startsWith("línea:")) {
          negocioActual = celdaC;
          i++; // Saltamos a la línea
          lineaActual = json[i][2]?.toString().trim();
          if (!dataCliente[negocioActual]) dataCliente[negocioActual] = {};
          if (!dataCliente[negocioActual][lineaActual]) dataCliente[negocioActual][lineaActual] = [];
          continue;
        }

        // Si celda C no es nula, pero no es línea ni negocio, es producto
        const esProducto = !esNegocio && !esLinea && celdaC !== undefined && celdaC !== null && celdaC !== "";

        if (esProducto && negocioActual && lineaActual) {
          const ventas = row.slice(12, 18).map(v => parseInt(v) || 0);
          const promedio = ventas.reduce((a, b) => a + b, 0) / ventas.length;

          dataCliente[negocioActual][lineaActual].push({
            producto: celdaC,
            ventasPorMes: ventas,
            etiquetas,
            promedio
          });
        }
      }

      clientsData[sheetName] = dataCliente;
    });

    renderClientButtons();
    showClient(Object.keys(clientsData)[0]);

    document.getElementById("btnExportarExcel").style.pointerEvents = "auto";
    document.getElementById("btnExportarExcel").style.opacity = "1";
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

  const selectedClientNameElement = document.getElementById("selectedClientName");
  selectedClientNameElement.textContent = `Cliente: ${clientName}`;
  selectedClientNameElement.style.color = "white";
  selectedClientNameElement.style.textShadow = "0 0 5px #000000, 0 0 10px #000000";

  const resumen = clientsData[clientName];

  Object.entries(resumen).forEach(([negocio, lineas]) => {
    const negocioBox = document.createElement("div");
    negocioBox.className = "negocio-container";
    negocioBox.style.margin = "20px auto";
    negocioBox.style.maxWidth = "1000px";

    const tituloNegocio = document.createElement("h3");
    tituloNegocio.textContent = negocio;
    tituloNegocio.style.color = "#0078d7";
    negocioBox.appendChild(tituloNegocio);

    const table = document.createElement("table");

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    // Encabezados: Línea | Producto | Promedio | Meses dinámicos
    ["Línea", "Producto", "Promedio"].forEach(h => {
      const th = document.createElement("th");
      th.textContent = h;
      headerRow.appendChild(th);
    });

    const exampleProducto = Object.values(lineas)[0]?.[0];
    const etiquetas = exampleProducto?.etiquetas || ["Mes 1", "Mes 2", "Mes 3", "Mes 4", "Mes 5", "Mes 6"];
    etiquetas.forEach(mes => {
      const th = document.createElement("th");
      th.textContent = mes;
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    Object.entries(lineas).forEach(([linea, productos]) => {
      // Calcular promedio real de todos los valores
      let suma = 0;
      let count = 0;

      productos.forEach(p => {
        p.ventasPorMes.forEach(v => {
          if (!isNaN(v)) {
            suma += v;
            count++;
          }
        });
      });

      const promedioLinea = count > 0 ? Math.round(suma / count) : 0;

      // Subtítulo línea
      const filaLinea = document.createElement("tr");
      const tdLinea = document.createElement("td");
      tdLinea.colSpan = 9;
      tdLinea.innerHTML = `
        <div class="linea-subtitulo" style="font-weight:bold; padding:10px 0 5px 0;">${linea}</div>
        <div style="margin: 5px 0;"><strong>Promedio total de ventas (línea):</strong> ${promedioLinea}</div>`;
      filaLinea.appendChild(tdLinea);
      tbody.appendChild(filaLinea);

      productos.forEach(prod => {
        const row = document.createElement("tr");

        const tdLineaNombre = document.createElement("td");
        tdLineaNombre.textContent = ""; // Línea ya está arriba

        const tdProd = document.createElement("td");
        tdProd.textContent = prod.producto;

        const tdProm = document.createElement("td");
        tdProm.textContent = Math.round(prod.promedio);

        row.appendChild(tdLineaNombre);
        row.appendChild(tdProd);
        row.appendChild(tdProm);

        prod.ventasPorMes.forEach(v => {
          const td = document.createElement("td");
          td.textContent = v;
          row.appendChild(td);
        });

        tbody.appendChild(row);
      });

      // Botón gráfico
      const filaBoton = document.createElement("tr");
      const tdBoton = document.createElement("td");
      tdBoton.colSpan = 9;
      tdBoton.style.textAlign = "right";

      const graficaBtn = document.createElement("button");
      graficaBtn.className = "toggle-btn";
      graficaBtn.textContent = "Gráfica de ventas";
      graficaBtn.onclick = () => mostrarGraficaLinea(linea, productos);

      tdBoton.appendChild(graficaBtn);
      filaBoton.appendChild(tdBoton);
      tbody.appendChild(filaBoton);
    });

    table.appendChild(tbody);
    negocioBox.appendChild(table);
    container.appendChild(negocioBox);
  });
}


function mostrarGraficaLinea(linea, productos) {
  if (!productos || productos.length === 0 || !productos[0].etiquetas) {
    alert("No hay datos disponibles para esta línea.");
    return;
  }

  const etiquetas = productos[0].etiquetas;

  const datasets = productos.map(prod => ({
    label: prod.producto,
    data: prod.ventasPorMes,
    borderColor: getRandomColor(),
    fill: false
  }));

  const ctx = document.getElementById("modalChartCanvas").getContext("2d");

  if (window.currentChartInstance) {
    window.currentChartInstance.destroy();
  }

  window.currentChartInstance = new Chart(ctx, {
    type: 'line',
    data: {
      labels: etiquetas,
      datasets
    },
    options: {
      responsive: true,
      plugins: {
        title: {
          display: true,
          text: `Gráfico de ventas por producto - ${linea}`,
          font: { size: 18 }
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

function exportarTodoExcel() {
  const wb = XLSX.utils.book_new();

  for (const cliente in clientsData) {
    const resumen = clientsData[cliente];
    const data = [];

    for (const negocio in resumen) {
      data.push([`NEGOCIO: ${negocio}`]);
      const lineas = resumen[negocio];
      for (const linea in lineas) {
        data.push([`  Línea: ${linea}`]);
        data.push(["Producto", "Promedio", "Ventas por Mes"]);
        lineas[linea].forEach(prod => {
          data.push([prod.producto, prod.promedio.toFixed(1), prod.ventasPorMes.join(", ")]);
        });
        data.push([]);
      }
      data.push([]);
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, cliente);
  }

  XLSX.writeFile(wb, "ResumenClientes.xlsx");
}
