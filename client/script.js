const yearSelect = document.querySelector("#yearSelect");
const dataDisplay = document.querySelector("#dataDisplay");
const dataForm = document.querySelector("#dataForm");

dataForm.addEventListener("submit", async (e) => {
  e.preventDefault();

  const selectedYear = yearSelect.value;

  const emissionsFile = "../data/emission.xlsx";
  const investmentsFile = "../data/investments.xlsx";

  try {
    const [emissionsData, investmentsData] = await Promise.all([
      loadExcelData(emissionsFile),
      loadExcelData(investmentsFile),
    ]);

    const filteredEmissions = filterDataByYear(emissionsData, selectedYear);
    const filteredInvestments = filterDataByYear(investmentsData, selectedYear);

    displayData(filteredEmissions, filteredInvestments);
  } catch (error) {
    console.error("Помилка завантаження файлів Excel:", error);
  }
});

const loadExcelData = async (filename) => {
  try {
    const response = await fetch(filename);
    if (!response.ok) {
      throw new Error("Помилка завантаження файлу");
    }
    const data = await response.arrayBuffer();
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet, { raw: true });
  } catch (error) {
    throw error;
  }
};

const filterDataByYear = (data, year) => {
  return data.filter((item) => item["Рік"] == year);
};

const displayData = (emissionsData, investmentsData) => {
  const emissionsTable = createTable(emissionsData, "Емісії");
  const investmentsTable = createTable(investmentsData, "Інвестиції");
  dataDisplay.innerHTML = "";
  dataDisplay.appendChild(emissionsTable);
  dataDisplay.appendChild(investmentsTable);
};

const createTable = (data, title) => {
  const table = document.createElement("table");
  table.innerHTML = `<caption>${title}</caption><thead></thead><tbody></tbody>`;
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");

  if (data.length === 0) {
    tbody.innerHTML = "<tr><td>Немає даних для відображення</td></tr>";
    return table;
  }

  const columns = Object.keys(data[0]);

  const headerRow = document.createElement("tr");
  columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  data.forEach((rowData) => {
    const row = document.createElement("tr");
    columns.forEach((column) => {
      const cell = document.createElement("td");
      cell.textContent = rowData[column];
      row.appendChild(cell);
    });
    tbody.appendChild(row);
  });

  return table;
};
