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
          const ventas = row.slice(12, 18).map(v => {
            const num = parseFloat(v);
            return isNaN(num) ? 0 : num;
          });

          const promedio = ventas.length > 0
            ? ventas.reduce((a, b) => a + b, 0) / ventas.length
            : 0;

          const inventario = parseFloat(row[19]) || 0;
          const pedidoSugerido = parseFloat(row[20]) || 0;

          dataCliente[negocioActual][lineaActual].push({
            producto: celdaC,
            promedio,
            inventario,
            pedidoSugerido
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
    negocioBox.style.width = "100%";
    negocioBox.style.maxWidth = "100%";

    const tituloNegocio = document.createElement("h3");
    tituloNegocio.textContent = negocio;
    tituloNegocio.style.color = "#0078d7";
    negocioBox.appendChild(tituloNegocio);

    const table = document.createElement("table");

    const thead = document.createElement("thead");
    const headerRow = document.createElement("tr");

    ["Línea", "Suma promedios", "Inventario + Venta", "Resultado"].forEach(h => {
      const th = document.createElement("th");
      th.textContent = h;
      headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement("tbody");

    Object.entries(lineas).forEach(([linea, productos]) => {
      let sumaPromedios = 0;
      let sumaInventario = 0;
      let sumaPedido = 0;

      productos.forEach(prod => {
        sumaPromedios += prod.promedio;
        sumaInventario += prod.inventario;
        sumaPedido += prod.pedidoSugerido;
      });

      const totalInvVentas = sumaInventario + sumaPedido;
      const resultado = sumaPromedios === 0 ? 0 : totalInvVentas / sumaPromedios;
      const promedioTotalLinea = productos.length > 0 ? sumaPromedios / productos.length : 0;

      const fila = document.createElement("tr");

      const tdLinea = document.createElement("td");
      tdLinea.innerHTML = `
        <div class="linea-subtitulo" style="font-weight:bold; padding:10px 0 5px 0;">${linea}</div>
        <div style="margin: 5px 0;">
          <strong>Promedio total de ventas (línea):</strong> ${Math.round(promedioTotalLinea)}
        </div>`;

      const tdSumaPromedios = document.createElement("td");
      tdSumaPromedios.textContent = Math.round(sumaPromedios);

      const tdInvVentas = document.createElement("td");
      tdInvVentas.textContent = Math.round(totalInvVentas);

      const tdResultado = document.createElement("td");
      tdResultado.textContent = resultado.toFixed(2).replace(".", ",");

      fila.appendChild(tdLinea);
      fila.appendChild(tdSumaPromedios);
      fila.appendChild(tdInvVentas);
      fila.appendChild(tdResultado);

      tbody.appendChild(fila);
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

document.getElementById("btnExportarExcel").addEventListener("click", exportarTodoExcel);

function exportarTodoExcel() {
  const wb = XLSX.utils.book_new();

  Object.entries(clientsData).forEach(([cliente, resumen]) => {
    const data = [];
    data.push(["Línea", "Promedio total de ventas (línea)", "Suma promedios", "Suma inv vent", "Resultado"]);

    Object.entries(resumen).forEach(([linea, productos]) => {
      let sumaPromedios = 0;
      let sumaInventario = 0;
      let sumaPedido = 0;

      productos.forEach(prod => {
        sumaPromedios += prod.promedio;
        sumaInventario += prod.inventario;
        sumaPedido += prod.pedidoSugerido;
      });

      const promedioLinea = productos.length > 0 ? Math.round(sumaPromedios / productos.length) : 0;
      const sumaInvVent = Math.round(sumaInventario + sumaPedido);
      const resultado = sumaPromedios === 0 ? 0 : (sumaInvVent / sumaPromedios);

      data.push([
        linea,
        promedioLinea,
        Math.round(sumaPromedios),
        sumaInvVent,
        resultado.toFixed(2).replace(".", ",")
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, cliente.substring(0, 31));
  });

  XLSX.writeFile(wb, "ResumenPorCliente.xlsx");
}
